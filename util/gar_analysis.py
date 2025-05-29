"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_analysis.py

    Author:       BCTS TOC - Graydon Shevchenko
    Purpose:      Tool used to run GAR and LRMP landbase analysis.
    Date Created: January 10, 2022
----------------------------------------------------------------------------------------------------------------
"""

# Import libraries
import gc
import traceback
import arcpy
import os
import sys
import logging

from argparse import ArgumentParser
from datetime import datetime as dt, timedelta
from collections import defaultdict

# Import classes
sys.path.insert(1, r'\\spatialfiles2.bcgov\work\FOR\RSI\TOC\Projects\ESRI_Scripts\Python_Repository')
sys.path.insert(2, r'\\spatialfiles2.bcgov\work\FOR\RSI\TOC\Projects\ESRI_Scripts\consolidated_cutblocks')
from environment import Environment
from create_consolidated_cutblocks import ConsolidatedCutblock
from util.gar_classes import GARInput, GARConfig, SICReplacement
from gar.gar_4001 import Gar4001
from gar.gar_4007 import Gar4007
from gar.gar_4010 import Gar4010
from gar.gar_8001 import Gar8001
from gar.gar_8005 import Gar8005
from gar.gar_8006 import Gar8006
from gar.gar_8012 import Gar8012
from gar.gar_8232 import Gar8232
from gar.lrmp_sheep import LrmpSheep


def run_app():
    """
    Function:
        Runs the main logic of the tool
    Returns:
        None
    """
    gar, out_gdb, out_fld, bec, run_cc, b_un, b_pw, logger = get_input_parameters()
    analysis = GARAnalysis(gar=gar, output_gdb=out_gdb, output_folder=out_fld, bec=bec,
                           run_cc=run_cc, bcgw_un=b_un, bcgw_pw=b_pw, logger=logger)
    analysis.prepare_data()
    analysis.identity_gar()
    if analysis.gar in ['u-8-006', 'u-8-001', 'u-8-001-tfl49']:
        analysis.add_sic_replacement()
    analysis.fix_slivers()
    analysis.calculate_values()
    analysis.dissolve_resultant()
    analysis.gar_class.write_excel()

    del analysis


def get_input_parameters():
    """
    Function:
        Sets up parameters and the logger object
    Returns:
        tuple: user entered parameters required for tool execution
    """
    try:
        parser = ArgumentParser(description='This script is used to analyze the landbase and report out on '
                                            'GAR and LRMP obligations')
        parser.add_argument('gar', type=str, help='GAR analysis to run')
        parser.add_argument('out_gdb', type=str, help='Output geodatabase')
        parser.add_argument('out_fld', type=str, help='Output folder location')
        parser.add_argument('bec', type=str, help='BEC Version', default='CURRENT', choices=['CURRENT', 'VERSION 5'])
        parser.add_argument('run_cc', type=str, help='Run consolidated cutblock')
        parser.add_argument('b_un', type=str, help='BCGW Username')
        parser.add_argument('b_pw', type=str, help='BCGW Password')
        parser.add_argument('--log_level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                            help='Log level')
        parser.add_argument('--log_dir', help='Path to log directory')

        args = parser.parse_args()

        logger = Environment.setup_logger(args)
        b_pw = arcpy.GetParameterAsText(6)

        return args.gar, args.out_gdb, args.out_fld, args.bec, args.run_cc, args.b_un, b_pw, logger

    except Exception as e:
        logging.error('Unexpected exception. Program terminating: {}'.format(e.message))
        raise Exception('Errors exist')


class GARAnalysis:
    """
    Class:
        GAR Analysis class containing methods for running the gar analysis
    """

    def __init__(self, gar, output_gdb, output_folder, bcgw_un, bcgw_pw, bec, run_cc, logger):
        """
        Function:
            Initializes the GARAnalysis class and all its attributes

        Args:
            gar (str): the gar analysis to run
            output_gdb (str): path to the output geodatabase
            output_folder (str): path to the output folder
            bcgw_un (str): username for the BCGW database
            bcgw_pw (str): password for the BCGW database
            bec (str): the BEC type to run in the analysis
            run_cc (str): string boolean indicating if consolidated cutblocks should be run
            logger (logger): logger object for writing messages to various output windows
        Returns:
            None
        """
        arcpy.env.overwriteOutput = True

        # Read in and assing input parameters
        self.gar = gar
        self.output_gdb = output_gdb
        self.output_fd = os.path.join(self.output_gdb, self.gar.replace('-', ''))
        self.output_folder = os.path.join(output_folder, self.gar.replace('-', '_'))
        if self.gar in ['lrmp-bhs', 'lrmp-ds']:
            self.output_xls = os.path.join(self.output_folder,
                                           'Report_LRMP_{0}_{1}_{2}_{3}.xlsx'.format(
                                               self.gar.replace('lrmp-bhs', 'Big_Horn_Sheep').replace('lrmp-ds',
                                                                                                      'Derenzy_Sheep'),
                                               dt.now().year, dt.now().month,
                                               dt.now().day))
        else:
            self.output_xls = os.path.join(self.output_folder,
                                           'Report_GAR_{0}_{1}_{2}_{3}.xlsx'.format(self.gar.replace('-', ''),
                                                                                    dt.now().year, dt.now().month,
                                                                                    dt.now().day))
        self.bcgw_un = bcgw_un
        self.bcgw_pw = bcgw_pw
        self.bec_version = bec
        self.run_cc = True if run_cc.lower() == 'true' else False
        self.logger = logger
        self.lrm_un = 'map_view_14'
        self.lrm_pw = 'interface'
        self.scratch_gdb = os.path.join(os.path.dirname(self.output_gdb), 'GAR_Scratch.gdb')
        self.sde_folder = output_folder
        self.cur_year = dt.now().year
        self.gar_class = None

        self.logger.info('Running analysis on {0}'.format(self.gar))
        # Connect to SDE databases and create output geodatabases
        self.lrm_db = Environment.create_lrm_connection(location=self.sde_folder, lrm_user_name=self.lrm_un,
                                                        lrm_password=self.lrm_pw, logger=self.logger)

        self.bcgw_db = Environment.create_bcgw_connection(location=self.sde_folder, bcgw_user_name=self.bcgw_un,
                                                          bcgw_password=self.bcgw_pw, logger=self.logger)

        if not arcpy.Exists(dataset=self.output_gdb):
            arcpy.CreateFileGDB_management(out_folder_path=os.path.dirname(self.output_gdb),
                                           out_name=os.path.basename(self.output_gdb))

        if not arcpy.Exists(dataset=self.output_fd):
            arcpy.CreateFeatureDataset_management(out_dataset_path=os.path.dirname(self.output_fd),
                                                  out_name=os.path.basename(self.output_fd),
                                                  spatial_reference=arcpy.SpatialReference(item=3005))

        # try:
        #     arcpy.Delete_management(in_data=self.scratch_gdb)
        # except (ValueError, Exception):
        #     pass
        if not arcpy.Exists(dataset=self.scratch_gdb):
            arcpy.CreateFileGDB_management(out_folder_path=os.path.dirname(self.scratch_gdb),
                                           out_name=os.path.basename(self.scratch_gdb))

        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

        # Dictionary of BEC layers including the lable field
        self.dict_bec = {
            'VERSION 5': [r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Misc_Data.gdb\BEC\BEC_v5', 'BECLABEL'],
            'CURRENT': [os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.BEC_BIOGEOCLIMATIC_POLY'), 'BGC_LABEL']
        }

        # Source Data

        # Change out operating areas from DBP06 to local data cut where feature called !non BCTS is added for u8-001 and u8-006 - Daniel Jan 12, 2025
        # if self.gar == 'u-8-001':
        #     self.__op_areas = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Misc_Data.gdb\Scripting_data\u_8001_Operating_Areas_4_Sciprting'
        # elif self.gar == 'u-8-001-tfl49':
        #     self.__op_areas = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Misc_Data.gdb\Scripting_data\u_8001_tfl_49_Operating_Areas_4_Sciprting'
        # elif self.gar == 'u-8-006':
        #     self.__op_areas = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Misc_Data.gdb\Scripting_data\u_8006_Operating_Areas_4_Sciprting'
        # else:
        self.__op_areas = os.path.join(self.lrm_db, 'BCTS_SPATIAL.BCTS_PROV_OP', 'BCTS_SPATIAL.OPERATING_AREA')

        self.__toc_area = os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.FADM_BCTS_AREA_SP')
        self.__uwr = os.path.join(self.bcgw_db, 'WHSE_WILDLIFE_MANAGEMENT.WCP_UNGULATE_WINTER_RANGE_SP')
        self.__uwr_golden = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Misc_Data.gdb\UWR\uwr_golden'
        self.__sec7 =r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Misc_Data.gdb\Golden_Sec_7_UWR\Mgmt_Unit_Boundaries'
        self.__wha = os.path.join(self.bcgw_db, 'WHSE_WILDLIFE_MANAGEMENT.WCP_WILDLIFE_HABITAT_AREA_POLY')
        self.__lrmp = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_PLAN_NON_LEGAL_POLY_SVW')
        self.__lrmp2 = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_PLAN_LEGAL_POLY_SVW')
        self.__lu = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_LANDSCAPE_UNIT_SP')
        self.__vri = os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.VEG_COMP_LYR_R1_POLY')
        self.__tfl = os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.FADM_TFL')
        self.__burn_severity = r'\\spatialfiles.bcgov\work\!Shared_Access\BARC\2024\Same_Year' \
                               r'\provincial_burn_severity_2024.gdb\provincial_burn_severity_2024'
        self.__fire_perimeters = os.path.join(self.bcgw_db, 'WHSE_LAND_AND_NATURAL_RESOURCE.PROT_CURRENT_FIRE_POLYS_SP')
        self.__fire_perimeters_hist = \
            os.path.join(self.bcgw_db, 'WHSE_LAND_AND_NATURAL_RESOURCE.PROT_HISTORICAL_FIRE_POLYS_SP')
        self.__bec = self.dict_bec[self.bec_version][0]
        self.__mot_roads = os.path.join(self.bcgw_db, 'WHSE_IMAGERY_AND_BASE_MAPS.MOT_ROAD_FEATURES_INVNTRY_SP')
        self.__ften_roads = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_ROAD_SECTION_LINES_SVW')
        self.__ften_blks = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_CUT_BLOCK_POLY_SVW')
        self.__results_inv = os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.RSLT_FOREST_COVER_INV_SVW')

        # This layer not available in BCGW anymore - Daniel Jan. 1, 2025
        #self.__private_land_lrdw = os.path.join(self.bcgw_db, 'WHSE_CADASTRE.CBM_INTGD_CADASTRAL_FABRIC_SVW')
        self.__private_land_pmbc = os.path.join(self.bcgw_db, 'WHSE_CADASTRE.PMBC_PARCEL_FABRIC_POLY_SVW')
        self.__woodlots = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_MANAGED_LICENCE_POLY_SVW')
        self.__slope = r'\\spatialfiles2.bcgov\Archive\FOR\RSI\TOC\Local_Data\Data_Library\terrain\Slope\Slope80.gdb' \
                       r'\Slope80_LiDAR_DEM_Merge_Single'

        # Per Michael Kyler the 8-001 should be using the old TSR2 THLB
        if self.gar == 'u-8-001':
            self.__thlb = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Historical_THLB.gdb\THLB_TSR2_OKTSA'
        elif self.gar == 'u-8-001-tfl49':
            self.__thlb = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\Historical_THLB.gdb\THLB_TFL49_Sept_2020'
        else:
            self.__thlb = r'\\spatialfiles2.bcgov\Archive\FOR\VIC\HTS\FAIB_DATA_FOR_DISTRIBUTION\THLB\Archive\TSA_THLB.gdb\prov_THLB_tsas'

        self.__consolidated_cb = r'\\spatialfiles2.bcgov\Archive\FOR\RSI\TOC\Local_Data\Data_Library\forest' \
                                 r'\consolidated_cutblocks\consolidated_cutblocks.gdb\ConsolidatedCutblocks_Prod_Res'
        self.__csrd_parks = r'\\spatialfiles2.bcgov\archive\FOR\RSI\TOC\Local_Data\Data_Library\Recreation\csrd_parks' \
                            r'\Parks.gdb\Parks'
        self.__prov_parks = os.path.join(self.bcgw_db, 'WHSE_TANTALIS.TA_PARK_ECORES_PA_SVW')
        self.__nat_parks = os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.CLAB_NATIONAL_PARKS')
        self.__crown_grants = os.path.join(self.bcgw_db, 'WHSE_LEGAL_ADMIN_BOUNDARIES.ILRR_LAND_ACT_CROWN_GRANTS_SVW')
        self.__xmas_tree_permits = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_HARVEST_AUTH_POLY_SVW')
        self.__sic_replacement = r'\\bctsdata.bcgov\data\toc_root\Genus_Reporting\GIS_spatial\SIC_Replacement' \
                                 r'\SIC_Replacement.gdb\Replacement_Areas'
        self.__CFLB_Selkirk = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\CFLB_THLB.gdb\CFLB_THLB\Selkirk_CFLB_4_mapping'
        self.__CFLB_Okanagan = r'\\bctsdata.bcgov\data\toc_root\Local_Data\Planning_data\CFLB_THLB.gdb\CFLB_THLB\Ok_CFLB_4_mapping'

        # Output Data
        self.fc_op_areas = os.path.join(self.scratch_gdb, 'op_areas')
        self.fc_toc_area = os.path.join(self.scratch_gdb, 'toc_area')
        self.fc_tfl49 = os.path.join(self.scratch_gdb, 'tfl49')
        self.fc_gar_cells = os.path.join(self.output_fd, '{}_UWR'.format(self.gar.replace('-', '')))
        self.fc_gar_cells_erase = os.path.join(self.scratch_gdb, 'gar_cells_erase')
        self.fc_lu = os.path.join(self.scratch_gdb, 'lu')
        self.fc_vri = os.path.join(self.scratch_gdb, 'vri')
        self.fc_vri_clip = os.path.join(self.scratch_gdb, 'vri_clip')
        self.fc_burn_severity = os.path.join(self.scratch_gdb, 'burn_severity')
        self.fc_fire_perimeters = os.path.join(self.scratch_gdb, 'fire_perimeters')
        self.fc_fire_perimeters_hist = os.path.join(self.scratch_gdb, 'fire_perimeters_hist')
        self.fc_bec = os.path.join(self.scratch_gdb, 'bec')
        self.fc_mot_roads = os.path.join(self.scratch_gdb, 'mot_roads')
        self.fc_ften_roads = os.path.join(self.scratch_gdb, 'ften_roads')
        self.fc_private_land = os.path.join(self.scratch_gdb, 'private_land')
        self.fc_federal_land = os.path.join(self.scratch_gdb, 'federal_land')
        self.fc_crown_grants = os.path.join(self.scratch_gdb, 'crown_grants')
        self.fc_csrd_parks = os.path.join(self.scratch_gdb, 'csrd_parks')
        self.fc_prov_parks = os.path.join(self.scratch_gdb, 'prov_parks')
        self.fc_nat_parks = os.path.join(self.scratch_gdb, 'nat_parks')
        self.fc_woodlots = os.path.join(self.scratch_gdb, 'woodlots')
        self.fc_slope = os.path.join(self.scratch_gdb, 'slope')
        self.fc_thlb = os.path.join(self.scratch_gdb, 'thlb')
        self.fc_xmas_trees = os.path.join(self.scratch_gdb, 'xmas_trees')
        self.fc_sic_replacement = os.path.join(self.scratch_gdb, 'sic_replacement')
        self.fc_consolidated_cb = os.path.join(self.scratch_gdb, 'blocks')
        self.fc_burn_areas = os.path.join(self.scratch_gdb, 'burn_areas')
        self.fc_broadleaf_stands = os.path.join(self.scratch_gdb, 'broadleaf_stands')
        self.fc_erase_features = os.path.join(self.scratch_gdb, 'erase_features')
        self.fc_road_merge = os.path.join(self.scratch_gdb, 'road_merge')
        self.fc_road_buffer = os.path.join(self.scratch_gdb, 'road_buffer')
        self.fc_road_dissolve = os.path.join(self.scratch_gdb, 'road_dissolve'.format(self.gar.replace('-', '')))
        self.fc_gar_cells_identity = os.path.join(self.scratch_gdb, 'gar_identity'.format(self.gar.replace('-', '')))
        self.fc_gar_cells_single = os.path.join(self.scratch_gdb, 'gar_single'.format(self.gar.replace('-', '')))
        self.fc_resultant = os.path.join(self.output_fd, '{}_Resultant'.format(self.gar.replace('-', '')))
        self.fc_resultant_dissolve = '{0}_Dissolve'.format(self.fc_resultant)
        self.fc_resultant_rank = os.path.join(self.output_fd, '{}_Resultant_Rank'.format(self.gar.replace('-', '')))
        self.fc_recent_ften_blks = os.path.join(self.scratch_gdb, 'recent_ften_blks')
        self.fc_results_res = os.path.join(self.scratch_gdb, 'results_reserves')

        # Dictionary of all inputs required for this analysis including selection criteria for creating a subset
        self.dict_gar_inputs = {
            'burn_severity': GARInput(path=self.__burn_severity, output=self.fc_burn_severity, mandatory=True),
            'fire_perimeters': GARInput(path=self.__fire_perimeters, output=self.fc_fire_perimeters, mandatory=True),
            'fire_perimeters_hist': GARInput(path=self.__fire_perimeters_hist, output=self.fc_fire_perimeters_hist,
                                             sql='FIRE_YEAR = {0}'.format(dt.now().year - 1),
                                             mandatory=True),
            'bec': GARInput(path=self.__bec, output=self.fc_bec, mandatory=True),
            'mot_roads': GARInput(path=self.__mot_roads, output=self.fc_mot_roads, mandatory=True),
            'ften_roads': GARInput(path=self.__ften_roads,
                                   sql='FILE_TYPE_DESCRIPTION IN(\'Forest Service Road\' , \'Road Permit\')',
                                   output=self.fc_ften_roads, mandatory=True),
            'private_land': GARInput(path=self.__private_land_pmbc, sql='OWNER_TYPE NOT IN (\'Crown Agency\' , '
                                                                        '\'Crown Provincial\' , \'Unclassified\' , '
                                                                        '\'Untitled Provincial\')',
                                     output=self.fc_private_land, mandatory=True),
            'woodlots': GARInput(path=self.__woodlots, sql='LIFE_CYCLE_STATUS_CODE = \'ACTIVE\'',
                                 output=self.fc_woodlots),
            'lu': GARInput(path=self.__lu, output=self.fc_lu),
            'slope': GARInput(path=self.__slope, output=self.fc_slope),
            'thlb': GARInput(path=self.__thlb, sql='tsa_number in (22,27,7,45) and thlb_fact > 0', output=self.fc_thlb),
            'blocks': GARInput(path=self.__consolidated_cb, output=self.fc_consolidated_cb),
            #'federal_land': GARInput(path=self.__private_land_lrdw, sql='OWNERSHIP_CLASS = \'CROWN FEDERAL\'',
                                     #output=self.fc_federal_land),
            'federal_land': GARInput(path=self.__private_land_pmbc, sql='OWNER_TYPE = \'Federal\' AND PARCEL_STATUS = \'Active\' AND PARCEL_CLASS = \'Crown Subdivision\'',
                                     output=self.fc_federal_land),
            'csrd_parks': GARInput(path=self.__csrd_parks, sql='ParkType = \'Community Park\'',
                                   output=self.fc_csrd_parks),
            'prov_parks': GARInput(path=self.__prov_parks, sql='PROTECTED_LANDS_CODE <> \'RC\'',
                                   output=self.fc_prov_parks),
            'nat_parks': GARInput(path=self.__nat_parks, output=self.fc_nat_parks),
            'crown_grants': GARInput(path=self.__crown_grants, output=self.fc_crown_grants),
            'xmas_trees': GARInput(path=self.__xmas_tree_permits,
                                   sql='LIFE_CYCLE_STATUS_CODE = \'ACTIVE\' AND FEATURE_CLASS_SKEY = 489',
                                   output=self.fc_xmas_trees),
            'recent_ften_blks': GARInput(path=self.__ften_blks, 
                                       sql="DISTURBANCE_START_DATE > TIMESTAMP '{0}'".format(
                                           (dt.now() - timedelta(days=5*365)).strftime('%Y-%m-%d %H:%M:%S')),
                                       output=self.fc_recent_ften_blks),
            'results_reserves': GARInput(path=self.__results_inv, sql ='(SILV_RESERVE_CODE = \'W\' or '
                                                                        'SILV_RESERVE_OBJECTIVE_CODE = \'WTR\') or '
                                                                        '(STOCKING_STATUS_CODE = \'MAT\' and STOCKING_TYPE_CODE = '
                                                                        '\'NAT\')\')', output=self.fc_results_res)
        }

        # Field names
        self.fld_line_7_activity = 'LINE_7_ACTIVITY_HIST_SYMBOL'
        self.fld_line_7b_dist_hist = 'LINE_7B_DISTURBANCE_HISTORY'
        self.fld_fire_version = 'VERSION_NUMBER'
        self.fld_burn_severity = 'BURN_SEVERITY_RATING'
        self.fld_fire_area = 'FIRE_SIZE_HECTARES'
        self.fld_fire_number = 'FIRE_NUMBER'
        self.fld_road_buffer = 'ROAD_BUFFER'
        self.fld_age_cur = 'AGE_CUR'
        self.fld_height_cur = 'HEIGHT_CUR'
        self.fld_height_text = 'HEIGHT_TEXT'
        self.fld_level = 'LEVEL'
        self.fld_rank_oa = 'OP_AREA_RANK'
        self.fld_rank_cell = 'CELL_RANK'
        self.fld_bec_version = 'BEC_VERSION'
        self.fld_date_created = 'DATE_CREATED'
        self.fld_crown_closure = 'CROWN_CLOSURE'
        self.fld_proj_date = 'PROJECTED_DATE'
        self.fld_proj_age = 'PROJ_AGE_1'
        self.fld_proj_height = 'PROJ_HEIGHT_1'
        self.fld_cc_status = 'CC_STATUS'
        self.fld_cc_harv_date = 'CC_HARVEST_DATE'
        self.fld_bec = self.dict_bec[self.bec_version][1]
        #Other BEC fields for use in u-8-232
        self.fld_bec_zone_alt = 'ZONE'
        self.fld_bec_subzone_alt = 'SUBZONE'
        self.fld_species = 'SPECIES_CD_1'
        self.fld_slope = 'SLOPE'
        self.fld_thlb = 'thlb_fact'
        self.fld_diameter = 'QUAD_DIAM_175'
        self.fld_percent = 'SPECIES_PCT_1'
        self.fld_notes = 'FEATURE_NOTES'
        self.fld_op_area = 'OPERATING_AREA'
        self.fld_shp_area = 'SHAPE@AREA'
        self.fld_calc_cflb = 'CALC_CFLB'
        self.fld_for_mgmt_ind = 'FOR_MGMT_LAND_BASE_IND'
        self.fld_bclcs_2 = 'BCLCS_LEVEL_2'
        self.fld_open_ind = 'OPENING_IND'
        self.fld_uwr_num = 'Name' if self.gar == 'section-7' else ('MGT' if self.gar == 'u-4-007' else 'UWR_UNIT_NUMBER')
        self.fld_bec_zone = 'BEC_ZONE_CODE'
        self.fld_bec_subzone = 'BEC_SUBZONE'
        self.fld_bec_variant = 'BEC_VARIANT'
        self.fld_lu = 'LANDSCAPE_UNIT_NAME'
        self.fld_lrmp = 'NON_LEGAL_FEAT_PROVID'
        self.fld_lrmp2 = 'LEGAL_FEAT_PROVID'
        self.fld_species_2 = 'SPECIES_CD_2'
        self.fld_species_3 = 'SPECIES_CD_3'
        self.fld_species_4 = 'SPECIES_CD_4'
        self.fld_species_5 = 'SPECIES_CD_5'
        self.fld_species_6 = 'SPECIES_CD_6'
        self.fld_percent_2 = 'SPECIES_PCT_2'
        self.fld_percent_3 = 'SPECIES_PCT_3'
        self.fld_percent_4 = 'SPECIES_PCT_4'
        self.fld_percent_5 = 'SPECIES_PCT_5'
        self.fld_percent_6 = 'SPECIES_PCT_6'

        # Set up the analysis configuration using the GarConfig class based on the input gar analysis to be run
        # Creates the applicable Gar class based on the selected gar
        if self.gar == 'u-4-001':
            gar_config = GARConfig(sql='UWR_NUMBER = \'{}\' AND FEATURE_NOTES NOT LIKE \'%SIC = 0%\''.format(self.gar),
                                   cells=self.__uwr,
                                   cell_field=self.fld_uwr_num,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_federal_land, self.fc_csrd_parks,
                                              self.fc_prov_parks, self.fc_nat_parks, self.fc_crown_grants,
                                              self.fc_broadleaf_stands],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = Gar4001(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-4-007':
            gar_config = GARConfig(cells=self.__uwr_golden,
                                   cell_field=self.fld_uwr_num,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_federal_land, self.fc_prov_parks,
                                              self.fc_nat_parks, self.fc_xmas_trees],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip, self.fc_slope]
                                   )
            self.gar_class = Gar4007(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-4-010':
            gar_config = GARConfig(sql='UWR_NUMBER = \'{}\' AND FEATURE_NOTES NOT LIKE \'%SIC = 0%\''.format(self.gar),
                                   cells=self.__uwr,
                                   cell_field=self.fld_notes,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_federal_land, self.fc_prov_parks,
                                              self.fc_nat_parks],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = Gar4010(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-8-001':
            gar_config = GARConfig(sql='UWR_NUMBER = \'{}\' AND FEATURE_NOTES NOT LIKE \'%SIC = 0%\''
                                   .format(self.gar.replace('-tfl49', '')),
                                   cells=self.__uwr,
                                   cell_field=self.fld_uwr_num,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_thlb, self.fc_vri_clip,
                                                 self.fc_slope]
                                   )
            self.gar_class = Gar8001(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-8-001-tfl49':
            gar_config = GARConfig(sql='UWR_NUMBER = \'{}\' AND FEATURE_NOTES NOT LIKE \'%SIC = 0%\''
                                   .format(self.gar.replace('-tfl49', '')),
                                   cells=self.__uwr,
                                   cell_field=self.fld_uwr_num,
                                   aoi=self.fc_tfl49,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_thlb, self.fc_vri_clip,
                                                 self.fc_slope]
                                   )
            self.gar_class = Gar8001(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-8-005':
            gar_config = GARConfig(sql='UWR_NUMBER = \'{}\' AND FEATURE_NOTES NOT LIKE \'%SIC = 0%\''.format(self.gar),
                                   cells=self.__uwr,
                                   cell_field=self.fld_uwr_num,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = Gar8005(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-8-006':
            gar_config = GARConfig(sql='UWR_NUMBER = \'{}\' AND FEATURE_NOTES NOT LIKE \'%SIC = 0%\''.format(self.gar),
                                   cells=self.__uwr,
                                   cell_field=self.fld_uwr_num,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = Gar8006(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-8-012':
            gar_config = GARConfig(sql='UWR_NUMBER = \'{}\' AND FEATURE_NOTES NOT LIKE \'%SIC = 0%\''.format(self.gar),
                                   cells=self.__uwr,
                                   cell_field=self.fld_bec,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = Gar8012(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'u-8-232':
            gar_config = GARConfig(sql='TAG = \'{}\' AND ORG_ORGANIZATION_ID IN (4, 8)'.format(self.gar[2:]),
                                   cells=self.__wha,
                                   cell_field=self.fld_lu,
                                   aoi=self.fc_op_areas,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots, self.fc_federal_land],
                                   identity_fcs=[self.fc_op_areas, self.fc_lu, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = Gar8232(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)
        elif self.gar == 'lrmp-bhs':
            gar_config = GARConfig(sql='STRGC_LAND_RSRCE_PLAN_NAME = \'Okanagan Shuswap Land and Resource Management '
                                       'Plan\' AND LEGAL_FEAT_OBJECTIVE = \'Big Horn Sheep Areas\'',
                                   cells=self.__lrmp2,
                                   cell_field=self.fld_lrmp2,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots, self.fc_federal_land],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = LrmpSheep(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                       gar_config=gar_config)
        elif self.gar == 'lrmp-ds':
            gar_config = GARConfig(sql='STRGC_LAND_RSRCE_PLAN_NAME = \'Okanagan Shuswap Land and Resource Management '
                                       'Plan\' AND NON_LEGAL_FEAT_OBJECTIVE = \'Derenzy Bighorn Sheep Habitat RMZ\' '
                                       'AND NON_LEGAL_FEAT_ATRB_1_VALUE = \'2\'',
                                   cells=self.__lrmp,
                                   cell_field=self.fld_lrmp,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = LrmpSheep(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                       gar_config=gar_config)
            
        elif self.gar == 'section-7':
            gar_config = GARConfig(cells=self.__sec7,
                                   cell_field=self.fld_uwr_num,
                                   aoi=self.fc_toc_area,
                                   private_land=self.__private_land_pmbc,
                                   erase_fcs=[self.fc_private_land, self.fc_woodlots],
                                   identity_fcs=[self.fc_op_areas, self.fc_bec, self.fc_road_dissolve,
                                                 self.fc_consolidated_cb, self.fc_vri_clip]
                                   )
            self.gar_class = Gar8006(gar=self.gar, output_xls=self.output_xls, logger=self.logger,
                                     gar_config=gar_config)            

    def __del__(self):
        """
        Function:
            Called when the class object is deleted; cleans up database connections and the scratch geodatabase
        Returns:
            None
        """
        Environment.delete_lrm_connection(location=self.sde_folder, logger=self.logger)
        Environment.delete_bcgw_connection(location=self.sde_folder, logger=self.logger)
        # arcpy.env.workspace = self.scratch_gdb
        # for f in arcpy.ListFeatureClasses():
        #     if f not in [self.fc_resultant, self.fc_gar_cells]:
        #         arcpy.Delete_management(os.path.join(self.scratch_gdb, f))

    def prepare_data(self):
        """
        Function:
            Prepares all required data for the analysis; the consolidated cutblock process
            will also be called from here if required
        Returns:
            None
        """
        if self.run_cc:
            # Call the consolidated cutblock process if the user has selected the option to run
            self.logger.info('Running consolidated cutblock subroutine')
            cc = ConsolidatedCutblock(bcgw_un=self.bcgw_un, bcgw_pw=self.bcgw_pw,
                                      output_gdb=os.path.dirname(self.__consolidated_cb), logger=self.logger)
            cc.prepare_data()
            cc.combine_data()
            cc.calculate_info()
            cc.flatten_fc()
            cleaned_fc = cc.check_geometry_area()
            self.logger.info('Make final feature class')
            final_consolidated = os.path.join(cc.output_gdb, 'ConsolidatedCutblocks_Prod_Res')
            arcpy.CopyFeatures_management(in_features=cleaned_fc, out_feature_class=final_consolidated)
            del cc
            self.logger.info('Completed consolidated cutblock subroutine')

        # Copying operating areas, toc boundary, and tfl49
        self.logger.info('Copying operating areas')
        arcpy.Select_analysis(in_features=self.__op_areas, out_feature_class=self.fc_op_areas,
                              where_clause='ORG_UNIT_CODE = \'TOC\'')
        self.logger.info('Copying bcts boundary')
        arcpy.Select_analysis(in_features=self.__toc_area, out_feature_class=self.fc_toc_area,
                              where_clause='BCTS_NAME = \'Okanagan - Columbia Timber Sales Business Area\'')
        self.logger.info('Copying tfl49')
        arcpy.Select_analysis(in_features=self.__tfl, out_feature_class=self.fc_tfl49,
                              where_clause='FOREST_FILE_ID = \'TFL49\'')
        oa_lyr = arcpy.MakeFeatureLayer_management(in_features=self.gar_class.gar_config.aoi, out_layer='oa_lyr')

        # Extracting the applicable cells for use as the aoi
        self.logger.info('Comparing to gar cells')
        gar_lyr = arcpy.MakeFeatureLayer_management(in_features=self.gar_class.gar_config.cells,
                                                    out_layer='gar_lyr',
                                                    where_clause=self.gar_class.gar_config.sql)

        arcpy.SelectLayerByLocation_management(in_layer=gar_lyr, overlap_type='INTERSECT', select_features=oa_lyr)
        arcpy.CopyFeatures_management(in_features=gar_lyr, out_feature_class=self.fc_gar_cells)

        arcpy.Delete_management(in_data=oa_lyr)
        arcpy.Delete_management(in_data=gar_lyr)

        # Merging cells together as needed
        self.merge_cells()

        # Set the processing extent
        self.logger.info('Setting extent')
        arcpy.env.extent = arcpy.Describe(value=self.fc_gar_cells).extent

        # Copy the vri and clip to the aoi
        self.logger.info('Copying vri')
        gar_lyr = arcpy.MakeFeatureLayer_management(in_features=self.fc_gar_cells, out_layer='gar_lyr')
        vri_lyr = arcpy.MakeFeatureLayer_management(in_features=self.__vri, out_layer='vri_lyr')
        arcpy.SelectLayerByLocation_management(in_layer=vri_lyr, overlap_type='INTERSECT', select_features=gar_lyr)
        arcpy.CopyFeatures_management(in_features=vri_lyr, out_feature_class=self.fc_vri)
        arcpy.Delete_management(in_data=vri_lyr)

        self.logger.info('Clipping vri to gar cells')
        arcpy.Clip_analysis(in_features=self.fc_vri, clip_features=self.fc_gar_cells,
                            out_feature_class=self.fc_vri_clip)

        # Copy the rest of the inputs, creating subsets where required
        for gar_input in self.dict_gar_inputs:
            if gar_input.startswith('private_land'):
                if self.dict_gar_inputs[gar_input].path != self.gar_class.gar_config.private_land:
                    continue
            if self.dict_gar_inputs[gar_input].mandatory or \
                    self.dict_gar_inputs[gar_input].output in self.gar_class.gar_config.erase_fcs or \
                    self.dict_gar_inputs[gar_input].output in self.gar_class.gar_config.identity_fcs:
                self.logger.info('Copying {0}'.format(gar_input))
                input_lyr = arcpy.MakeFeatureLayer_management(in_features=self.dict_gar_inputs[gar_input].path,
                                                              out_layer='input_lyr',
                                                              where_clause=self.dict_gar_inputs[gar_input].sql)
                arcpy.SelectLayerByLocation_management(in_layer=input_lyr, overlap_type='INTERSECT',
                                                       select_features=gar_lyr)
                arcpy.CopyFeatures_management(in_features=input_lyr,
                                              out_feature_class=self.dict_gar_inputs[gar_input].output)
                arcpy.Delete_management(in_data=input_lyr)

        arcpy.Delete_management(in_data=gar_lyr)

        # Add in burn severity if required for the selected gar
        # if self.fc_burn_areas in self.gar_class.gar_config.erase_fcs or \
        #         self.fc_burn_areas in self.gar_class.gar_config.identity_fcs:
        self.add_burn_severity()

        # Add in broadleaf stands if required for the selected gar
        if self.fc_broadleaf_stands in self.gar_class.gar_config.erase_fcs or \
                self.fc_broadleaf_stands in self.gar_class.gar_config.identity_fcs:
            self.create_broadleaf_stand_layer()

        # Removing all required features from the aoi
        self.logger.info('Combining features to erase')
        arcpy.Merge_management(inputs=self.gar_class.gar_config.erase_fcs, output=self.fc_erase_features)

        self.logger.info('Erasing features from gar cells')
        arcpy.Erase_analysis(in_features=self.fc_gar_cells, erase_features=self.fc_erase_features,
                             out_feature_class=self.fc_gar_cells_erase)

        # Creating the road right of ways
        self.logger.info('Creating road right of ways')
        arcpy.Merge_management(inputs=[self.fc_mot_roads, self.fc_ften_roads], output=self.fc_road_merge)
        arcpy.Buffer_analysis(in_features=self.fc_road_merge, out_feature_class=self.fc_road_buffer,
                              buffer_distance_or_field='10 Meters', dissolve_option='NONE')
        arcpy.AddField_management(in_table=self.fc_road_buffer, field_name=self.fld_road_buffer,
                                  field_type='TEXT', field_length=3)
        with arcpy.da.UpdateCursor(self.fc_road_buffer, self.fld_road_buffer) as u_cursor:
            for row in u_cursor:
                row[0] = 'YES'
                u_cursor.updateRow(row)
        arcpy.Dissolve_management(in_features=self.fc_road_buffer, out_feature_class=self.fc_road_dissolve,
                                  dissolve_field=self.fld_road_buffer, multi_part='SINGLE_PART')
        
        #Creating Recent Harvest Area to account for harvest areas not yet populated by VRI - added by Daniel Otto March 24, 2025
        if self.gar == 'section-7':
            self.logger.info('Creating Recent Harvest Area')

    def add_burn_severity(self):
        """
        Function:
            Compares the vri, fire perimeters and burn severity layer to create the burned areas
            required for removal from the analysis area.  These areas are not deleted, instead the age and height are
            reduced to zero with the projected date being updated to the fire year
        Returns:
            None
        """
        # Appending the previous years historical fires to the current years fires
        self.logger.info('Adding burn severity to vri')
        arcpy.Append_management(inputs=self.fc_fire_perimeters_hist, target=self.fc_fire_perimeters,
                                schema_type='NO_TEST')

        # If there are no fires found in both this year and last year, then return
        if int(arcpy.GetCount_management(in_rows=self.fc_fire_perimeters).getOutput(0)) == 0:
            self.logger.warning('No fires found within area of interest')
            # self.gar_class.gar_config.erase_fcs.remove(self.fc_burn_areas)
            return

        fire_bs = os.path.join(self.scratch_gdb, 'fire_burn')
        vri_burn = os.path.join(self.scratch_gdb, 'vri_burn')
        lst_fields = [self.fld_line_7_activity, self.fld_line_7b_dist_hist, self.fld_fire_version,
                      self.fld_burn_severity, self.fld_fire_area, self.fld_fire_number, 'SHAPE@', self.fld_proj_age,
                      self.fld_proj_height, self.fld_proj_date]

        # Combine the fire perimeters, burn severity and vri
        arcpy.Identity_analysis(in_features=self.fc_fire_perimeters, identity_features=self.fc_burn_severity,
                                out_feature_class=fire_bs)
        arcpy.Identity_analysis(in_features=self.fc_vri_clip, identity_features=fire_bs, out_feature_class=vri_burn)

        # Loop through the resultant and find areas that need age and height adjusted to zero
        with arcpy.da.UpdateCursor(vri_burn, lst_fields) as u_cursor:
            for row in u_cursor:
                line_7 = row[lst_fields.index(self.fld_line_7_activity)]
                line_7b = str(row[lst_fields.index(self.fld_line_7b_dist_hist)])
                fire_version = row[lst_fields.index(self.fld_fire_version)]
                fire_burn_severity = row[lst_fields.index(self.fld_burn_severity)]
                fire_area = row[lst_fields.index(self.fld_fire_area)]
                fire_year = int(str(fire_version)[:4])
                fire_num = row[lst_fields.index(self.fld_fire_number)]
                proj_date = row[lst_fields.index(self.fld_proj_date)]
                if fire_num == '':  # If there is no fire number, skip this row
                    continue
                if line_7 == '$':  # If there is a disturbance identified in the vri layer
                    if line_7b.startswith('B'):  # If the disturbance is a fire
                        dist_year = line_7b[-2:]
                        # If the disturbance year is the same as the fire year,
                        # then it's already accounted for in the vri, skip this row
                        if dist_year == str(fire_year)[-2:]:
                            continue
                if fire_area < 100:  # If the fire area is less than 100 hectares
                    # If there is no burn severity value, then assume it's High and adjust age, height and date
                    if not fire_burn_severity:
                        row[lst_fields.index(self.fld_proj_age)] = 0
                        row[lst_fields.index(self.fld_proj_height)] = 0
                        new_date = proj_date.replace(year=fire_year)
                        row[lst_fields.index(self.fld_proj_date)] = new_date
                        u_cursor.updateRow(row)
                        continue
                # If the burn severity is Medium or High adjust age, height and date
                if fire_burn_severity in ['Medium', 'High']:
                    row[lst_fields.index(self.fld_proj_age)] = 0
                    row[lst_fields.index(self.fld_proj_height)] = 0
                    new_date = proj_date.replace(year=fire_year)
                    row[lst_fields.index(self.fld_proj_date)] = new_date
                    u_cursor.updateRow(row)
                    continue

        arcpy.CopyFeatures_management(in_features=vri_burn, out_feature_class=self.fc_vri_clip)
        arcpy.Delete_management(in_data=fire_bs)
        arcpy.Delete_management(in_data=vri_burn)

    def add_sic_replacement(self):
        fld_bec_zone = 'BEC_ZONE_CODE'
        fld_bec_subzone = 'BEC_SUBZONE'
        fld_bec_var = 'BEC_VARIANT'
        fld_age = 'AGE'
        fld_dbh = 'DBH'
        fld_height = 'HEIGHT'
        fld_crown = 'CROWN_CLOSURE'
        fld_slope = 'SLOPE'
        fld_spec1 = 'SPECIES_1'
        fld_perc1 = 'SPECIES_PCT_1'
        fld_spec2 = 'SPECIES_2'
        fld_perc2 = 'SPECIES_PCT_2'
        fld_spec3 = 'SPECIES_3'
        fld_perc3 = 'SPECIES_PCT_3'
        fld_spec4 = 'SPECIES_4'
        fld_perc4 = 'SPECIES_PCT_4'
        fld_spec5 = 'SPECIES_5'
        fld_perc5 = 'SPECIES_PCT_5'
        fld_spec6 = 'SPECIES_6'
        fld_perc6 = 'SPECIES_PCT_6'
        fld_survey_date = 'SURVEY_DATE'

        dict_replacement = defaultdict(SICReplacement)

        lst_fields = ['OID@',fld_bec_zone, fld_bec_subzone, fld_bec_var, fld_age, fld_dbh, fld_height, fld_crown,
                      fld_slope, fld_spec1, fld_perc1, fld_spec2, fld_perc2, fld_spec3, fld_perc3, fld_spec4, fld_perc4,
                      fld_spec5, fld_perc5, fld_spec6, fld_perc6, fld_survey_date]

        self.logger.info('Copying SIC replacement areas')
        arcpy.CopyFeatures_management(in_features=self.__sic_replacement, out_feature_class=self.fc_sic_replacement)

        self.logger.info('Reading in replacement values')
        with arcpy.da.SearchCursor(self.fc_sic_replacement, lst_fields) as s_cursor:
            for row in s_cursor:
                oid = row[lst_fields.index('OID@')]
                dict_replacement[oid].zone = row[lst_fields.index(fld_bec_zone)]
                dict_replacement[oid].sub = row[lst_fields.index(fld_bec_subzone)]
                dict_replacement[oid].var = row[lst_fields.index(fld_bec_var)]
                dict_replacement[oid].age = row[lst_fields.index(fld_age)]
                dict_replacement[oid].dbh = row[lst_fields.index(fld_dbh)]
                dict_replacement[oid].hgt = row[lst_fields.index(fld_height)]
                dict_replacement[oid].cc = row[lst_fields.index(fld_crown)]
                dict_replacement[oid].slp = '80+' if row[lst_fields.index(fld_slope)] >= 80 else None
                dict_replacement[oid].sp1 = row[lst_fields.index(fld_spec1)]
                dict_replacement[oid].per1 = row[lst_fields.index(fld_perc1)]
                dict_replacement[oid].sp2 = row[lst_fields.index(fld_spec2)]
                dict_replacement[oid].per2 = row[lst_fields.index(fld_perc2)]
                dict_replacement[oid].sp3 = row[lst_fields.index(fld_spec3)]
                dict_replacement[oid].per3 = row[lst_fields.index(fld_perc3)]
                dict_replacement[oid].sp4 = row[lst_fields.index(fld_spec4)]
                dict_replacement[oid].per4 = row[lst_fields.index(fld_perc4)]
                dict_replacement[oid].sp5 = row[lst_fields.index(fld_spec5)]
                dict_replacement[oid].per5 = row[lst_fields.index(fld_perc5)]
                dict_replacement[oid].sp6 = row[lst_fields.index(fld_spec6)]
                dict_replacement[oid].per6 = row[lst_fields.index(fld_perc6)]
                dict_replacement[oid].survey_dt = row[lst_fields.index(fld_survey_date)]

        vri_sic = os.path.join(self.scratch_gdb, 'vri_sic')
        arcpy.Identity_analysis(in_features=self.fc_gar_cells_identity, identity_features=self.fc_sic_replacement,
                                out_feature_class=vri_sic, join_attributes='ONLY_FID')
        fld_oid = 'FID_sic_replacement'
        lst_fields = [fld_oid, self.fld_bec, self.fld_proj_age,
                      self.fld_diameter, self.fld_proj_height, self.fld_crown_closure, self.fld_species,
                      self.fld_percent, self.fld_species_2, self.fld_percent_2,
                      self.fld_species_3, self.fld_percent_3, self.fld_species_4, self.fld_percent_4,
                      self.fld_species_5, self.fld_percent_5, self.fld_species_6, self.fld_percent_6,
                      self.fld_proj_date]

        if self.fld_slope in [field.name for field in arcpy.ListFields(vri_sic)]:
            lst_fields.append(self.fld_slope)

        self.logger.info('Replacing values in vri')
        where_clause = '{0} <> -1'.format(fld_oid)
        with arcpy.da.UpdateCursor(vri_sic, lst_fields, where_clause=where_clause) as u_cursor:
            for row in u_cursor:
                oid = row[lst_fields.index(fld_oid)]
                row[lst_fields.index(self.fld_bec)] = \
                    '{0} {1} {2}'.format(dict_replacement[oid].zone, dict_replacement[oid].sub,
                                         dict_replacement[oid].var)
                row[lst_fields.index(self.fld_proj_age)] = dict_replacement[oid].age
                row[lst_fields.index(self.fld_diameter)] = dict_replacement[oid].dbh
                row[lst_fields.index(self.fld_proj_height)] = dict_replacement[oid].hgt
                row[lst_fields.index(self.fld_crown_closure)] = dict_replacement[oid].cc
                if self.fld_slope in lst_fields:
                    row[lst_fields.index(self.fld_slope)] = dict_replacement[oid].slp
                row[lst_fields.index(self.fld_species)] = dict_replacement[oid].sp1
                row[lst_fields.index(self.fld_percent)] = dict_replacement[oid].per1
                row[lst_fields.index(self.fld_species_2)] = dict_replacement[oid].sp2
                row[lst_fields.index(self.fld_percent_2)] = dict_replacement[oid].per2
                row[lst_fields.index(self.fld_species_3)] = dict_replacement[oid].sp3
                row[lst_fields.index(self.fld_percent_3)] = dict_replacement[oid].per3
                row[lst_fields.index(self.fld_species_4)] = dict_replacement[oid].sp4
                row[lst_fields.index(self.fld_percent_4)] = dict_replacement[oid].per4
                row[lst_fields.index(self.fld_species_5)] = dict_replacement[oid].sp5
                row[lst_fields.index(self.fld_percent_5)] = dict_replacement[oid].per5
                row[lst_fields.index(self.fld_species_6)] = dict_replacement[oid].sp6
                row[lst_fields.index(self.fld_percent_6)] = dict_replacement[oid].per6
                row[lst_fields.index(self.fld_proj_date)] = dict_replacement[oid].survey_dt
                u_cursor.updateRow(row)

        arcpy.CopyFeatures_management(in_features=vri_sic, out_feature_class=self.fc_gar_cells_identity)
        arcpy.Delete_management(in_data=vri_sic)

    def merge_cells(self):
        """
        Function:
            Dissolves all input cells and calculates the original cell id values
            into a new field for each dissolved feature
        Returns:
            None
        """
        if 'lrmp' not in self.gar:
            return

        # Dissolving gar cells into singlepart features and add in id field
        self.logger.info('Merging cells')
        temp_fc = os.path.join(self.scratch_gdb, 'temp_fc')
        arcpy.Dissolve_management(in_features=self.fc_gar_cells, out_feature_class=temp_fc, multi_part='SINGLE_PART')
        arcpy.AddField_management(in_table=temp_fc, field_name=self.gar_class.gar_config.cell_field,
                                  field_type='TEXT', field_length=200)


        arcpy.MultipartToSinglepart_management(in_features=self.fc_gar_cells, out_feature_class='singlepart_fc')

        # Create a new dictionary
        temp_dict = {}

        # Iterate over the features
        with arcpy.da.SearchCursor('singlepart_fc', [self.gar_class.gar_config.cell_field, 'SHAPE@']) as s_cursor:
            for row in s_cursor:
                # Check the area of the feature
                if row[1].getArea('PLANAR', 'SQUAREMETERS') >= 1000:
                    # If the area is 1000 or more, add it to the dictionary
                    temp_dict[row[0]] = row[1]


        # Loop through the records in the dissolved feature class creating strings of ids if they overlap
        # with the shapes in the original gar cell dictionary
        with arcpy.da.UpdateCursor(temp_fc, [self.gar_class.gar_config.cell_field, 'SHAPE@']) as u_cursor:
            for row in u_cursor:
                if row[1].getArea('PLANAR', 'SQUAREMETERS') < 1000:
                    # If the area is less than 1000, delete the feature
                    u_cursor.deleteRow()
                else:
                    str_id = ''
                    arcpy.AddMessage("Before checking contains: " + str_id)  # Print str_id before the loop
                    for obj in temp_dict.keys():
                        if row[1].contains(temp_dict[obj]):
                            str_id += '{}, '.format(obj)
                    arcpy.AddMessage("After checking contains: " + str_id)  # Print str_id after the loop
                    row[0] = str_id[:-2] if str_id != '' else str_id
                    u_cursor.updateRow(row)

        # Overwrite the input gar cells with the dissolved features
        arcpy.CopyFeatures_management(in_features=temp_fc, out_feature_class=self.fc_gar_cells)

    def create_broadleaf_stand_layer(self):
        """
        Function:
            Creates a broadleaf stand layer
        Returns:
            None
        """
        # Copy the vri
        fld_broadleaf_percent = 'Broadleaf_Percent'
        fld_bclcs_4 = 'BCLCS_LEVEL_4'
        arcpy.CopyFeatures_management(self.fc_vri_clip, self.fc_broadleaf_stands)
        arcpy.AddField_management(self.fc_broadleaf_stands, fld_broadleaf_percent, 'SHORT')

        layer_list = [fld_bclcs_4, fld_broadleaf_percent]
        for i in range(1, 7):
            layer_list.append('SPECIES_CD_{0}'.format(i))
            layer_list.append('SPECIES_PCT_{0}'.format(i))

        with arcpy.da.UpdateCursor(self.fc_broadleaf_stands, layer_list) as u_cursor:
            for row in u_cursor:
                land_class = row[layer_list.index('BCLCS_LEVEL_4')]

                if land_class == 'TB':
                    # do nothing
                    pass
                elif land_class == 'TM':
                    # TM land class must be over 50% of deciduous leading
                    percent = 0
                    for num in range(1, 7):
                        spp = row[layer_list.index('SPECIES_CD_' + str(num))]
                        pct = row[layer_list.index('SPECIES_PCT_' + str(num))]
                        if spp:
                            if spp.startswith('A') or spp.startswith('E') or spp in ['DR', 'MB']:
                                percent += pct
                    if percent <= 50:
                        u_cursor.deleteRow()
                    else:
                        row[layer_list.index(fld_broadleaf_percent)] = percent
                        u_cursor.updateRow(row)
                else:
                    u_cursor.deleteRow()

    def identity_gar(self):
        """
        Function:
            Runs the identity process with back up subroutines in the event the identity fails due to memory exceptions
        Returns:
            None
        """
        input_fc = self.fc_gar_cells_erase
        output_fc = os.path.join(self.scratch_gdb, 'temp_output')
        temp_input = os.path.join(self.scratch_gdb, 'temp_input')
        dice_temp = os.path.join(self.scratch_gdb, 'dice_temp')
        subdivide_poly = os.path.join(self.scratch_gdb, 'subdivide_poly')

        # Loop through the identity features in the configuration for the gar
        for ident_lyr in self.gar_class.gar_config.identity_fcs:
            self.logger.info('Adding {0} to gar cells'.format(os.path.basename(ident_lyr)))
            try:
                # Try running the Identity tool
                arcpy.Identity_analysis(in_features=input_fc, identity_features=ident_lyr,
                                        out_feature_class=output_fc, join_attributes='NO_FID')

            except (ValueError, Exception):
                try:
                    # If the Identity tool fails (often due to memory issues) run Dice on the features and try again
                    self.logger.warning('...File too large, dicing')
                    arcpy.Dice_management(in_features=ident_lyr, out_feature_class=dice_temp, vertex_limit=10000)
                    self.logger.info('...Attempting identity again')
                    arcpy.Identity_analysis(in_features=input_fc, identity_features=dice_temp,
                                            out_feature_class=output_fc, join_attributes='NO_FID')
                except (ValueError, Exception):
                    try:
                        # If the Identity fails again, subdivide the polygons, dice, then identity again
                        self.logger.warning('...File contains overly large polygons, subdividing')
                        subdivide_poly = Environment.subdivide_polygons(input_fc=ident_lyr, output_fc=subdivide_poly)
                        self.logger.info('...Attempting dice again')
                        arcpy.Dice_management(in_features=subdivide_poly, out_feature_class=dice_temp,
                                              vertex_limit=10000)

                        self.logger.info('...Attempting identity again')
                        arcpy.Identity_analysis(in_features=input_fc, identity_features=dice_temp,
                                                out_feature_class=output_fc, join_attributes='NO_FID')
                    except (ValueError, Exception):
                        self.logger.error(traceback.format_exc())
                        return
            arcpy.CopyFeatures_management(in_features=output_fc, out_feature_class=temp_input)
            input_fc = temp_input

        arcpy.CopyFeatures_management(in_features=output_fc, out_feature_class=self.fc_gar_cells_identity)
        for lyr in [dice_temp, subdivide_poly, output_fc, temp_input]:
            if arcpy.Exists(lyr):
                arcpy.Delete_management(in_data=lyr)

    def fix_slivers(self):
        """
        Function:
            Cleans up the resultant dataset by converting to singlepart, repairing geometry and removing slivers
        Returns:
            None
        """
        single_part_output = self.fc_gar_cells_single
        fld_area = 'Area_m'

        # Convert to singlepart and repair geometry
        self.logger.info('Converting to single part')
        arcpy.MultipartToSinglepart_management(in_features=self.fc_gar_cells_identity,
                                               out_feature_class=single_part_output)
        self.logger.info('Repairing geometry')
        arcpy.RepairGeometry_management(in_features=single_part_output)

        # Update the area field with the updated shape areas
        arcpy.AddField_management(in_table=single_part_output, field_name=fld_area, field_type='DOUBLE')
        with arcpy.da.UpdateCursor(single_part_output, ['SHAPE@AREA', fld_area]) as u_cursor:
            for row in u_cursor:
                row[1] = row[0]
                u_cursor.updateRow(row)
        prev_selection = 9999999999
        output_fc = os.path.join(self.scratch_gdb, 'out_temp')
        output_temp_fc = output_fc

        # Run eliminate polygons for the first time
        current_selection = self.eliminate_small_polygons(inputfc=single_part_output, outputfc=output_fc,
                                                          area_field=fld_area)

        self.logger.info('Merge 1m polygons with biggest neighbour')
        # do eliminates until there are no more polygons neighbouring to join to.
        while prev_selection > current_selection:
            self.logger.info('{} polygon(s) remaining'.format(current_selection))
            # run it again:
            input_fc = output_fc
            output_temp_fc = os.path.join(self.scratch_gdb, 'out_temp_1')
            prev_selection = current_selection
            # Run eliminate polygons
            current_selection = self.eliminate_small_polygons(inputfc=input_fc, outputfc=output_temp_fc,
                                                              area_field=fld_area)

        # Once all slivers have beeen eliminated create resultant
        arcpy.DeleteField_management(in_table=output_temp_fc, drop_field=fld_area)
        self.logger.info('Creating resultant')
        arcpy.CopyFeatures_management(in_features=output_temp_fc, out_feature_class=self.fc_resultant)
        for f in [output_fc, output_temp_fc, single_part_output]:
            if arcpy.Exists(dataset=f):
                arcpy.Delete_management(in_data=f)

    def eliminate_small_polygons(self, inputfc, outputfc, area_field):
        """
        Function:
            Loops through the input feature class and eliminates slivers
        Args:
            inputfc (str): path to the input feature class
            outputfc (str): path to the output feature class
            area_field (str): area field used for determining feature areas

        Returns:
            int: number of features remaining that are still considered slivers
        """
        # Select all polygons that are less than 1 square metre
        temp_layer = arcpy.MakeFeatureLayer_management(in_features=inputfc, out_layer='temp_lyr')
        arcpy.SelectLayerByAttribute_management(in_layer_or_view=temp_layer, selection_type='NEW_SELECTION',
                                                where_clause='{0} < 1'.format(area_field))
        current_selection = int(arcpy.GetCount_management(in_rows=temp_layer).getOutput(0))
        gc.collect()
        arcpy.Delete_management(in_data='in_memory')
        try:
            # Run eliminate on the selected polygons
            arcpy.Eliminate_management(in_features=temp_layer, out_feature_class=outputfc, selection='AREA')
        except:
            # If eliminate fails, run eliminate on each gar cell instead and merge to create final file
            self.logger.warning('Eliminate failed due to large dataset size, '
                                'running eliminate based on {0}'.format(self.gar_class.gar_config.cell_field))
            temp_fc = os.path.join(self.scratch_gdb, 'eliminate_temp')
            lst_ids = sorted(list(set([row[0] for row in
                                       arcpy.da.SearchCursor(inputfc, self.gar_class.gar_config.cell_field)])))
            if arcpy.Exists(dataset=outputfc):
                arcpy.Delete_management(in_data=outputfc)
            for cell_id in lst_ids:
                self.logger.info('Working on {0}'.format(cell_id))
                temp_layer = \
                    arcpy.MakeFeatureLayer_management(in_features=inputfc, out_layer='temp_lyr',
                                                      where_clause='{0} = \'{1}\''.format(
                                                          self.gar_class.gar_config.cell_field, cell_id))
                arcpy.SelectLayerByAttribute_management(in_layer_or_view=temp_layer, selection_type='NEW_SELECTION',
                                                        where_clause='{0} < 1'.format(area_field))
                arcpy.Eliminate_management(in_features=temp_layer, out_feature_class=temp_fc, selection='AREA')
                if not arcpy.Exists(dataset=outputfc):
                    arcpy.CopyFeatures_management(in_features=temp_fc, out_feature_class=outputfc)
                else:
                    arcpy.Append_management(inputs=temp_fc, target=outputfc, schema_type='NO_TEST')

        # Update area field with new shape area
        with arcpy.da.UpdateCursor(outputfc, ['SHAPE@AREA', area_field]) as u_cursor:
            for row in u_cursor:
                row[1] = row[0]
                u_cursor.updateRow(row)

        return current_selection

    def calculate_values(self):
        """
        Function:
            Adds new fields and calculates values based off the attributes found in the
            resultant data; targets are also calculated in this function
        Returns:
            None
        """
        # Add required fields
        for fld in [self.fld_age_cur, self.fld_height_cur, self.fld_height_text, self.fld_level, self.fld_rank_oa,
                    self.fld_rank_cell, self.fld_bec_version, self.fld_date_created, self.fld_calc_cflb]:
            field_type = 'TEXT'
            if fld in self.fld_age_cur:
                field_type = 'SHORT'
            elif fld == self.fld_date_created:
                field_type = 'DATE'
            elif fld == self.fld_height_cur:
                field_type = 'DOUBLE'
            try:
                arcpy.AddField_management(in_table=self.fc_resultant, field_name=fld, field_type=field_type)
            except (ValueError, Exception):
                pass

        self.logger.info('Updating age and collecting areas')
        current_year = dt.now().year
        field_list = [self.fld_proj_date, self.fld_proj_age, self.fld_age_cur, self.fld_road_buffer, self.fld_cc_status,
                      self.fld_cc_harv_date, self.fld_bec_version, self.fld_date_created, self.fld_bec, self.fld_level,
                      self.fld_species, self.fld_crown_closure, self.fld_slope, self.fld_thlb, self.fld_diameter,
                      self.fld_percent, self.fld_notes, self.fld_op_area, self.fld_shp_area, self.fld_calc_cflb,
                      self.fld_bclcs_2, self.fld_open_ind, self.fld_line_7b_dist_hist,
                      self.gar_class.gar_config.cell_field, self.fld_proj_height, self.fld_height_cur,
                      self.fld_height_text, self.fld_for_mgmt_ind]
        field_list = [f for f in field_list if f in
                      [field.name for field in arcpy.ListFields(dataset=self.fc_resultant)] or f == self.fld_shp_area]
        # Loop through resultant calculating values needed for this analysis
        with arcpy.da.UpdateCursor(self.fc_resultant, field_list) as u_cursor:
            for row in u_cursor:
                # Read in values from resultant record
                proj_date = row[field_list.index(self.fld_proj_date)]
                proj_age = row[field_list.index(self.fld_proj_age)]
                proj_hgt = row[field_list.index(self.fld_proj_height)]
                rd_buffer = row[field_list.index(self.fld_road_buffer)]
                cc_status = row[field_list.index(self.fld_cc_status)]
                cc_harv_date = row[field_list.index(self.fld_cc_harv_date)]
                bec = str(row[field_list.index(self.fld_bec)]).replace(' ', '')
                spp = str(row[field_list.index(self.fld_species)])
                cc = row[field_list.index(self.fld_crown_closure)]
                slp = row[field_list.index(self.fld_slope)] if self.fld_slope in field_list else None
                thlb = (float(row[field_list.index(self.fld_thlb)]) if row[field_list.index(self.fld_thlb)] else 0) \
                    if self.fld_thlb in field_list else None
                diam = row[field_list.index(self.fld_diameter)]
                pct = row[field_list.index(self.fld_percent)]
                notes = row[field_list.index(self.fld_notes)] if self.fld_notes in field_list else ''
                target = int(notes[notes.find('=') + 2:]) \
                    if any(char.isdigit() for char in notes) and '=' in notes else None
                pcell = row[field_list.index(self.gar_class.gar_config.cell_field)]
                op_area = row[field_list.index(self.fld_op_area)]
                shp_area = row[field_list.index(self.fld_shp_area)] / 10000
                for_ind = row[field_list.index(self.fld_for_mgmt_ind)]
                calc_cflb = None
                height_cur = None
                height_text = None
                age_cur = None

                if proj_date:  # If a date exists in the record
                    difference = current_year - proj_date.year
                    try:
                        # Add the difference in years to the existing age to make it current
                        age_cur = int(proj_age) + difference
                    except (ValueError, Exception):
                        pass

                    if proj_hgt:  # If a height exists in the record
                        # Grow the height by 30 centimetres per year
                        height_cur = int(proj_hgt) + (0.3 * difference)
                        if height_cur >= 19.5:
                            height_text = '>= 19.5m'

                if cc_harv_date != '' and cc_status not in ('ROAD', 'RESERVE'):  # If the record has been harvested
                    try:
                        # Update age to be the difference between harvest date and now
                        age_cur = current_year - int(cc_harv_date[0:4])
                    except (ValueError, Exception):
                        pass

                if cc_status == 'ROAD':  # Find road buffered and set the age to none
                    row[field_list.index(self.fld_road_buffer)] = 'Yes'
                    age_cur = None

                if rd_buffer == 'Yes':
                    age_cur = None

                if for_ind == 'Y':  # Check if the polygon is part of the cflb
                    calc_cflb = 'Y'
                    row[field_list.index(self.fld_calc_cflb)] = calc_cflb

                if self.gar != 'u-8-232':  # Run the gar class level calculation if the gar is not 8-232
                    level = self.gar_class.calculate_level(bec=bec, age=age_cur, spp=spp, cc=cc, slp=slp, thlb=thlb,
                                                           diam=diam, pct=pct, gfa=calc_cflb, notes=notes,
                                                           op_area=op_area, pcell=pcell, shp_area=shp_area,
                                                           target=target, height=height_cur)
                    row[field_list.index(self.fld_level)] = level

                # Update row records
                row[field_list.index(self.fld_age_cur)] = age_cur
                row[field_list.index(self.fld_height_cur)] = height_cur
                row[field_list.index(self.fld_height_text)] = height_text
                row[field_list.index(self.fld_bec_version)] = self.bec_version
                row[field_list.index(self.fld_date_created)] = dt.now().strftime('%d/%m/%Y')

                u_cursor.updateRow(row)

        if self.gar == 'u-8-232':  # Run level caluclation for gar 8-232 as its different than all others
            lst_fields = [self.fld_op_area, self.fld_lu, self.fld_bec_zone_alt, self.fld_bec_subzone_alt,
                          self.fld_level, self.fld_height_text]
            arcpy.Dissolve_management(in_features=self.fc_resultant, out_feature_class=self.fc_resultant_dissolve,
                                      dissolve_field=lst_fields)
            lst_fields.append('SHAPE@AREA')
            with arcpy.da.UpdateCursor(self.fc_resultant_dissolve, lst_fields) as u_cursor:
                for row in u_cursor:
                    hgt = row[lst_fields.index(self.fld_height_text)]
                    shp_area = row[lst_fields.index('SHAPE@AREA')] / 10000
                    bec = '{0} {1}'.format(row[lst_fields.index(self.fld_bec_zone_alt)],
                                           row[lst_fields.index(self.fld_bec_subzone_alt)])
                    op_area = row[lst_fields.index(self.fld_op_area)]
                    lu = row[lst_fields.index(self.fld_lu)]

                    level = self.gar_class.calculate_level(op_area=op_area, pcell='{0}: {1}'.format(lu, bec),
                                                           shp_area=shp_area, height=hgt)
                    row[lst_fields.index(self.fld_level)] = level
                    u_cursor.updateRow(row)

        # Calculate targets
        self.gar_class.calculate_targets()

        # If ranks were calculated, update the resultant with the ranks
        if self.gar_class.gar_config.ranks:
            self.logger.info('Updating resultant with ranks')
            field_list = [self.gar_class.gar_config.cell_field, self.fld_level, self.fld_op_area, self.fld_rank_oa,
                          self.fld_rank_cell, self.fld_bec]
            with arcpy.da.UpdateCursor(self.fc_resultant, field_list) as u_cursor:
                for row in u_cursor:
                    pcell = row[field_list.index(self.gar_class.gar_config.cell_field)]
                    level = str(row[field_list.index(self.fld_level)])
                    op_area = row[field_list.index(self.fld_op_area)]
                    bec = str(row[field_list.index(self.fld_bec)]).replace(' ', '')
                    if self.gar == 'u-8-006':
                        oa_rank = self.gar_class.dict_total_area[op_area].pcell[pcell].level[level].rank
                        cell_rank = self.gar_class.dict_cell_area[pcell].level[level].rank
                    else:
                        oa_rank = self.gar_class.dict_total_area[op_area].pcell[pcell].level[level].bec[bec].rank
                        cell_rank = self.gar_class.dict_cell_area[pcell].level[level].bec[bec].rank
                    row[field_list.index(self.fld_rank_oa)] = oa_rank
                    row[field_list.index(self.fld_rank_cell)] = cell_rank
                    u_cursor.updateRow(row)

            if self.gar == 'u-8-006':  # Calculate mature stands for 8-006
                self.logger.info('Calculating mature stands')
                for fld in [self.fld_rank_oa, self.fld_rank_cell]:
                    sql_all = '{0} IN (\'CH\', \'NH\')'.format(fld)
                    sql_mature = '{0} = \'Mature Cover\''.format(self.fld_level)
                    lst_fields = [self.gar_class.gar_config.cell_field, self.fld_op_area] \
                        if fld == self.fld_rank_oa else [self.gar_class.gar_config.cell_field]
                    self.calculate_mature_stands(where_clause=sql_all, dissolve_fields=lst_fields, run_type='All')
                    self.calculate_mature_stands(where_clause=sql_mature, dissolve_fields=lst_fields, run_type='Mature')

    def calculate_mature_stands(self, where_clause, dissolve_fields, run_type):
        """
        Function:
            Determines which stands in the VRI are considered mature
        Args:
            where_clause (str): sql clause used as a where statement to select out applicable recored
            dissolve_fields (list): list of fields used in the dissolve
            run_type (str): 'Mature' or 'All'

        Returns:
            None
        """
        # Select subset and dissolve
        result_lyr = arcpy.MakeFeatureLayer_management(in_features=self.fc_resultant, out_layer='result_lyr',
                                                       where_clause=where_clause)
        fc_dissolve = os.path.join(self.scratch_gdb, 'dissolve_temp')
        arcpy.Dissolve_management(in_features=result_lyr, out_feature_class=fc_dissolve,
                                  dissolve_field=dissolve_fields, multi_part='SINGLE_PART')
        arcpy.Delete_management(in_data=result_lyr)
        lst_fields = dissolve_fields + ['SHAPE@AREA']

        if self.fld_op_area in lst_fields:  # If operating area based, loop through and calculate mature stands
            for row in arcpy.da.SearchCursor(fc_dissolve, lst_fields):
                shp = row[lst_fields.index('SHAPE@AREA')] / 10000
                pcell = row[lst_fields.index(self.gar_class.gar_config.cell_field)]
                op_area = row[lst_fields.index(self.fld_op_area)]
                if shp >= 20:
                    if run_type == 'Mature':
                        self.gar_class.dict_total_area[op_area].pcell[pcell]. \
                            level[self.gar_class.str_mature].stand_hectares += shp
                    else:
                        self.gar_class.dict_total_area[op_area].pcell[pcell].stand_hectares += shp
        else:  # If planning cell based, loop through and calculate mature stands
            for row in arcpy.da.SearchCursor(fc_dissolve, lst_fields):
                shp = row[lst_fields.index('SHAPE@AREA')] / 10000
                pcell = row[lst_fields.index(self.gar_class.gar_config.cell_field)]
                if shp >= 20:
                    if run_type == 'Mature':
                        self.gar_class.dict_cell_area[pcell].level[self.gar_class.str_mature].stand_hectares += shp
                    else:
                        self.gar_class.dict_cell_area[pcell].stand_hectares += shp
        arcpy.Delete_management(in_data=fc_dissolve)

    def dissolve_resultant(self):
        """
        Function:
            Creates the dissolved resultant feature class
        Returns:
            None
        """
        self.logger.info('Dissolving resultant')
        lst_fields = [self.fld_uwr_num, self.fld_notes, self.fld_op_area, self.fld_level, self.fld_rank_cell,
                      self.fld_rank_oa, self.fld_bec, self.fld_bec_version, self.fld_date_created]
        lst_fields = [f for f in lst_fields if f in
                      [field.name for field in arcpy.ListFields(dataset=self.fc_resultant)]]
        arcpy.Dissolve_management(in_features=self.fc_resultant, out_feature_class=self.fc_resultant_rank,
                                  dissolve_field=lst_fields, multi_part='SINGLE_PART')


if __name__ == '__main__':
    run_app()

