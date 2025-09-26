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
from environment import Environment

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
    Runs the main logic of the tool (BCGW-only, no ConsolidatedCutblock).
    Expects get_input_parameters() to return:
        gar, out_gdb, out_fld, bec, aoi_fc, b_un, b_pw, logger
    """
    gar, out_gdb, out_fld, bec, aoi_fc, b_un, b_pw, logger = get_input_parameters()

    analysis = GARAnalysis(
        gar=gar,
        output_gdb=out_gdb,
        output_folder=out_fld,
        bec=bec,
        bcgw_un=b_un,
        bcgw_pw=b_pw,
        logger=logger,
        aoi=aoi_fc  # NEW: optional AOI for small/fast test runs
    )

    logger.info(f"Starting GAR analysis: {gar}")

    try:
        analysis.prepare_data()
        analysis.identity_gar()

        # Optional & guarded: SIC replacement relies on non-BCGW sources in many setups.
        if gar in ['u-8-006', 'u-8-001', 'u-8-001-tfl49'] and hasattr(analysis, 'add_sic_replacement'):
            try:
                analysis.add_sic_replacement()
            except Exception as e:
                logger.warning(f"Skipping SIC replacement (not available in SA build?): {e}")

        analysis.fix_slivers()
        analysis.calculate_values()
        analysis.dissolve_resultant()

        # Optional & guarded: only call if implemented by the selected GAR class
        if getattr(analysis, "gar_class", None) and hasattr(analysis.gar_class, "write_excel"):
            try:
                analysis.gar_class.write_excel()
            except Exception as e:
                logger.warning(f"Excel export skipped: {e}")

    finally:
        # Ensures any connection cleanup in __del__ is executed
        del analysis



def get_input_parameters():
    """
    Sets up parameters and the logger object.

    Returns:
        tuple: (gar, out_gdb, out_fld, bec, aoi_fc, b_un, b_pw, logger)

    ArcGIS Pro Script Tool parameter order (recommended):
      0 gar        (String)
      1 out_gdb    (Workspace)
      2 out_fld    (Folder)
      3 bec        (String; use 'CURRENT')
      4 aoi_fc     (Optional Polygon feature class/layer)
      5 b_un       (String; BCGW username)
      6 b_pw       (Password; BCGW password)
      7 log_level  (Optional String; DEBUG/INFO/WARNING/ERROR)
      8 log_dir    (Optional Folder)
    """
    try:
        # --- ArcGIS Pro Script Tool mode ---
        if arcpy.GetArgumentCount() > 0:
            gar        = arcpy.GetParameterAsText(0)
            out_gdb    = arcpy.GetParameterAsText(1)
            out_fld    = arcpy.GetParameterAsText(2)
            bec        = arcpy.GetParameterAsText(3) or "CURRENT"
            aoi_fc     = arcpy.GetParameterAsText(4) or None
            b_un       = arcpy.GetParameterAsText(5)
            b_pw       = arcpy.GetParameterAsText(6)
            log_level  = arcpy.GetParameterAsText(7) or "INFO"
            log_dir    = arcpy.GetParameterAsText(8) or None

            if bec != "CURRENT":
                raise ValueError("This simplified build is BCGW-only; set BEC to 'CURRENT'.")

            # Build a tiny args-like object for Environment.setup_logger(...)
            class _Args: pass
            _a = _Args()
            _a.log_level = log_level
            _a.log_dir   = log_dir
            logger = Environment.setup_logger(_a)

            return gar, out_gdb, out_fld, bec, aoi_fc, b_un, b_pw, logger

        # --- CLI mode ---
        parser = ArgumentParser(
            description="Analyze landbase and report on GAR/LRMP (simplified, BCGW-only)."
        )
        parser.add_argument("gar", type=str, help="GAR analysis to run (e.g., u-8-005)")
        parser.add_argument("out_gdb", type=str, help="Output file geodatabase")
        parser.add_argument("out_fld", type=str, help="Output/working folder")
        parser.add_argument("--bec", default="CURRENT", choices=["CURRENT"],
                            help="BEC Version (BCGW only: CURRENT)")
        parser.add_argument("--aoi", dest="aoi_fc", help="Optional AOI polygon FC/layer path")
        parser.add_argument("--bcgw_user", dest="b_un", required=True, help="BCGW Username")
        parser.add_argument("--bcgw_pw",   dest="b_pw", required=True, help="BCGW Password")
        parser.add_argument("--log_level", default="INFO",
                            choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="Log level")
        parser.add_argument("--log_dir", help="Path to log directory")

        args = parser.parse_args()
        logger = Environment.setup_logger(args)

        return args.gar, args.out_gdb, args.out_fld, args.bec, args.aoi_fc, args.b_un, args.b_pw, logger

    except Exception as e:
        logging.error(f"Unexpected exception. Program terminating: {str(e)}")
        raise



class GARAnalysis:
    """
    GAR Analysis class containing methods for running the gar analysis
    """

    def __init__(self, gar, output_gdb, output_folder, bcgw_un, bcgw_pw, bec, logger, aoi=None):
        """
        Initializes the GARAnalysis class and all its attributes

        Args:
            gar (str): the gar analysis to run
            output_gdb (str): path to the output geodatabase
            output_folder (str): path to the output folder
            bcgw_un (str): username for the BCGW database
            bcgw_pw (str): password for the BCGW database
            bec (str): the BEC type to run in the analysis (use 'CURRENT' for BCGW-only)
            logger (logger): logger object
            aoi (str|None): optional AOI polygon feature class/layer to limit processing
        """
        arcpy.env.overwriteOutput = True

        # Inputs & paths
        self.gar = gar
        self.output_gdb = output_gdb
        self.output_fd = os.path.join(self.output_gdb, self.gar.replace('-', ''))
        self.output_folder = os.path.join(output_folder, self.gar.replace('-', '_'))

        if self.gar in ['lrmp-bhs', 'lrmp-ds']:
            self.output_xls = os.path.join(
                self.output_folder,
                'Report_LRMP_{0}_{1}_{2}_{3}.xlsx'.format(
                    self.gar.replace('lrmp-bhs', 'Big_Horn_Sheep').replace('lrmp-ds', 'Derenzy_Sheep'),
                    dt.now().year, dt.now().month, dt.now().day
                )
            )
        else:
            self.output_xls = os.path.join(
                self.output_folder,
                'Report_GAR_{0}_{1}_{2}_{3}.xlsx'.format(
                    self.gar.replace('-', ''), dt.now().year, dt.now().month, dt.now().day
                )
            )

        self.bcgw_un = bcgw_un
        self.bcgw_pw = bcgw_pw
        self.bec_version = bec
        self.logger = logger
        self.aoi = aoi  # NEW: optional AOI
        self.scratch_gdb = os.path.join(os.path.dirname(self.output_gdb), 'GAR_Scratch.gdb')
        self.sde_folder = output_folder
        self.cur_year = dt.now().year
        self.gar_class = None

        self.logger.info('Running analysis on {0}'.format(self.gar))

        # --- BCGW connection only (no LRM) ---
        self.bcgw_db = Environment.create_bcgw_connection(
            location=self.sde_folder,
            bcgw_user_name=self.bcgw_un,
            bcgw_password=self.bcgw_pw,
            logger=self.logger
        )

        # --- Ensure workspaces exist ---
        if not arcpy.Exists(self.output_gdb):
            arcpy.CreateFileGDB_management(
                out_folder_path=os.path.dirname(self.output_gdb),
                out_name=os.path.basename(self.output_gdb)
            )

        if not arcpy.Exists(self.output_fd):
            arcpy.CreateFeatureDataset_management(
                out_dataset_path=os.path.dirname(self.output_fd),
                out_name=os.path.basename(self.output_fd),
                spatial_reference=arcpy.SpatialReference(3005)
            )

        if not arcpy.Exists(self.scratch_gdb):
            arcpy.CreateFileGDB_management(
                out_folder_path=os.path.dirname(self.scratch_gdb),
                out_name=os.path.basename(self.scratch_gdb)
            )

        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

        # ---------------- BCGW-only sources & labels ----------------
        # BEC layer + its label field
        self.dict_bec = {
            'CURRENT': [
                os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.BEC_BIOGEOCLIMATIC_POLY'),
                'BGC_LABEL'
            ]
        }
        self.__bec = self.dict_bec[self.bec_version][0]

        # ---------------- Core BCGW feature classes -----------------
        # Cell / planning overlays
        self.__uwr   = os.path.join(self.bcgw_db, 'WHSE_WILDLIFE_MANAGEMENT.WCP_UNGULATE_WINTER_RANGE_SP')
        self.__wha   = os.path.join(self.bcgw_db, 'WHSE_WILDLIFE_MANAGEMENT.WCP_WILDLIFE_HABITAT_AREA_POLY')
        self.__lrmp  = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_PLAN_NON_LEGAL_POLY_SVW')
        self.__lrmp2 = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_PLAN_LEGAL_POLY_SVW')
        self.__lu    = os.path.join(self.bcgw_db, 'WHSE_LAND_USE_PLANNING.RMP_LANDSCAPE_UNIT_SP')
        self.__tfl   = os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.FADM_TFL')  # used for TFL49 cases

        # Forest inventory & results/tenure
        self.__vri         = os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.VEG_COMP_LYR_R1_POLY')
        self.__results_inv = os.path.join(self.bcgw_db, 'WHSE_FOREST_VEGETATION.RSLT_FOREST_COVER_INV_SVW')
        self.__ften_roads  = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_ROAD_SECTION_LINES_SVW')
        self.__ften_blks   = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_CUT_BLOCK_POLY_SVW')
        self.__woodlots    = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_MANAGED_LICENCE_POLY_SVW')

        # Base mapping / admin / parks
        self.__mot_roads   = os.path.join(self.bcgw_db, 'WHSE_IMAGERY_AND_BASE_MAPS.MOT_ROAD_FEATURES_INVNTRY_SP')
        self.__private_land_pmbc = os.path.join(self.bcgw_db, 'WHSE_CADASTRE.PMBC_PARCEL_FABRIC_POLY_SVW')
        self.__prov_parks  = os.path.join(self.bcgw_db, 'WHSE_TANTALIS.TA_PARK_ECORES_PA_SVW')
        self.__nat_parks   = os.path.join(self.bcgw_db, 'WHSE_ADMIN_BOUNDARIES.CLAB_NATIONAL_PARKS')
        self.__crown_grants = os.path.join(self.bcgw_db, 'WHSE_LEGAL_ADMIN_BOUNDARIES.ILRR_LAND_ACT_CROWN_GRANTS_SVW')
        self.__xmas_tree_permits = os.path.join(self.bcgw_db, 'WHSE_FOREST_TENURE.FTEN_HARVEST_AUTH_POLY_SVW')

        # Fire perimeters (BCGW). Note: burn severity product was non-BCGW → removed.
        self.__fire_perimeters      = os.path.join(self.bcgw_db, 'WHSE_LAND_AND_NATURAL_RESOURCE.PROT_CURRENT_FIRE_POLYS_SP')
        self.__fire_perimeters_hist = os.path.join(self.bcgw_db, 'WHSE_LAND_AND_NATURAL_RESOURCE.PROT_HISTORICAL_FIRE_POLYS_SP')

        # ---------------- Removed (non-BCGW / LRM / local paths) ----------------
        # self.__op_areas (LRM), self.__toc_area (BCTS area), self.__uwr_golden, self.__sec7
        # self.__burn_severity (local BARC), self.__slope (local), THLB local variants, consolidated cutblocks
        # self.__csrd_parks (local), self.__sic_replacement (local), CFLB (local)


        #--------------------------------------------------------------------------------------------------------------------------------------------------
        # ---------------- Output Data (BCGW-only + AOI support) ----------------
        # AOI staging (we'll prepare these if user supplies an AOI)
        #self.fc_aoi_work   = os.path.join(self.scratch_gdb, 'aoi_work')
        #self.fc_aoi_3005   = os.path.join(self.scratch_gdb, 'aoi_3005')
        #self.fc_aoi_single = os.path.join(self.scratch_gdb, 'aoi_single')
        self.fc_aoi_clean  = os.path.join(self.scratch_gdb, 'aoi_clean')  # ← use this instead of fc_toc_area

        # TFL (used for -tfl49 variants, if applicable)
        self.fc_tfl49 = os.path.join(self.scratch_gdb, 'tfl49')

        # Selected cells (UWR/WHA/LRMP) and working copies
        self.fc_gar_cells        = os.path.join(self.output_fd, f"{self.gar.replace('-', '')}_UWR")
        self.fc_gar_cells_erase  = os.path.join(self.scratch_gdb, 'gar_cells_erase')

        # BCGW layers clipped to AOI/cells
        self.fc_lu               = os.path.join(self.scratch_gdb, 'lu')
        self.fc_vri              = os.path.join(self.scratch_gdb, 'vri')
        self.fc_vri_clip         = os.path.join(self.scratch_gdb, 'vri_clip')
        self.fc_fire_perimeters  = os.path.join(self.scratch_gdb, 'fire_perimeters')
        self.fc_fire_perimeters_hist = os.path.join(self.scratch_gdb, 'fire_perimeters_hist')
        self.fc_bec              = os.path.join(self.scratch_gdb, 'bec')
        self.fc_mot_roads        = os.path.join(self.scratch_gdb, 'mot_roads')
        self.fc_ften_roads       = os.path.join(self.scratch_gdb, 'ften_roads')
        self.fc_private_land     = os.path.join(self.scratch_gdb, 'private_land')
        self.fc_crown_grants     = os.path.join(self.scratch_gdb, 'crown_grants')
        self.fc_prov_parks       = os.path.join(self.scratch_gdb, 'prov_parks')
        self.fc_nat_parks        = os.path.join(self.scratch_gdb, 'nat_parks')
        self.fc_woodlots         = os.path.join(self.scratch_gdb, 'woodlots')
        self.fc_xmas_trees       = os.path.join(self.scratch_gdb, 'xmas_trees')

        # Combine/remove features, roads, identity chain
        self.fc_erase_features   = os.path.join(self.scratch_gdb, 'erase_features')
        self.fc_road_merge       = os.path.join(self.scratch_gdb, 'road_merge')
        self.fc_road_buffer      = os.path.join(self.scratch_gdb, 'road_buffer')
        self.fc_road_dissolve    = os.path.join(self.scratch_gdb, 'road_dissolve')
        self.fc_gar_cells_identity = os.path.join(self.scratch_gdb, 'gar_identity')
        self.fc_gar_cells_single = os.path.join(self.scratch_gdb, 'gar_single')

        # Results
        self.fc_resultant           = os.path.join(self.output_fd, f"{self.gar.replace('-', '')}_Resultant")
        self.fc_resultant_dissolve  = f"{self.fc_resultant}_Dissolve"
        self.fc_resultant_rank      = os.path.join(self.output_fd, f"{self.gar.replace('-', '')}_Resultant_Rank")

        # Optional/derived subsets (still BCGW-based)
        self.fc_recent_ften_blks = os.path.join(self.scratch_gdb, 'recent_ften_blks')
        self.fc_results_res      = os.path.join(self.scratch_gdb, 'results_reserves')
        #--------------------------------------------------------------------------------------------------------------------------------------------------



        # Dictionary of all inputs required for this analysis including selection criteria for creating a subset
        # Dictionary of inputs (BCGW-only) with optional SQL filters
        self.dict_gar_inputs = {
            # Fire perimeters (kept, but not mandatory since we dropped burn-severity logic)
            'fire_perimeters': GARInput(
                path=self.__fire_perimeters,
                output=self.fc_fire_perimeters
            ),
            'fire_perimeters_hist': GARInput(
                path=self.__fire_perimeters_hist,
                sql='FIRE_YEAR = {0}'.format(dt.now().year - 1),
                output=self.fc_fire_perimeters_hist
            ),

            # Core layers
            'bec': GARInput(
                path=self.__bec,
                output=self.fc_bec,
                mandatory=True
            ),
            'mot_roads': GARInput(
                path=self.__mot_roads,
                output=self.fc_mot_roads,
                mandatory=True
            ),
            'ften_roads': GARInput(
                path=self.__ften_roads,
                sql="FILE_TYPE_DESCRIPTION IN('Forest Service Road','Road Permit')",
                output=self.fc_ften_roads,
                mandatory=True
            ),
            'private_land': GARInput(
                path=self.__private_land_pmbc,
                sql=(
                    "OWNER_TYPE NOT IN ('Crown Agency','Crown Provincial','Unclassified','Untitled Provincial')"
                ),
                output=self.fc_private_land,
                mandatory=True
            ),

            # Useful but optional (copied when referenced in erase_fcs/identity_fcs for a given GAR)
            'woodlots': GARInput(
                path=self.__woodlots,
                sql="LIFE_CYCLE_STATUS_CODE = 'ACTIVE'",
                output=self.fc_woodlots
            ),
            'lu': GARInput(
                path=self.__lu,
                output=self.fc_lu
            ),
            'prov_parks': GARInput(
                path=self.__prov_parks,
                sql="PROTECTED_LANDS_CODE <> 'RC'",
                output=self.fc_prov_parks
            ),
            'nat_parks': GARInput(
                path=self.__nat_parks,
                output=self.fc_nat_parks
            ),
            'crown_grants': GARInput(
                path=self.__crown_grants,
                output=self.fc_crown_grants
            ),
            'xmas_trees': GARInput(
                path=self.__xmas_tree_permits,
                sql="LIFE_CYCLE_STATUS_CODE = 'ACTIVE' AND FEATURE_CLASS_SKEY = 489",
                output=self.fc_xmas_trees
            ),

            # Recent FTEN blocks (5 years)
            'recent_ften_blks': GARInput(
                path=self.__ften_blks,
                sql="DISTURBANCE_START_DATE > TIMESTAMP '{0}'".format(
                    (dt.now() - timedelta(days=5*365)).strftime('%Y-%m-%d %H:%M:%S')
                ),
                output=self.fc_recent_ften_blks
            ),

            # RESULTS reserves (typo fixed vs original extra quote/paren)
            'results_reserves': GARInput(
                path=self.__results_inv,
                sql=(
                    "(SILV_RESERVE_CODE = 'W' OR SILV_RESERVE_OBJECTIVE_CODE = 'WTR') OR "
                    "(STOCKING_STATUS_CODE = 'MAT' AND STOCKING_TYPE_CODE = 'NAT')"
                ),
                output=self.fc_results_res
            ),
        }



        #--------------------------------------------------------------------------------------------------------------------------------------------------

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




        #--------------------------------------------------------------------------------------------------------------------------------------------------

        # Set up the analysis configuration using the GarConfig class based on the input gar analysis to be run
        # Creates the applicable Gar class based on the selected gar
        if self.gar == 'u-4-001':
            gar_config = GARConfig(
                sql=f"UWR_NUMBER = '{self.gar}' AND FEATURE_NOTES NOT LIKE '%SIC = 0%'",
                cells=self.__uwr,
                cell_field=self.fld_uwr_num,
                aoi=self.fc_aoi_clean,  # was fc_toc_area (BCTS boundary); now your AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,   # PMBC parcels (non-Crown etc.)
                    self.fc_woodlots,       # ACTIVE woodlots
                    self.fc_prov_parks,     # provincial parks (excluding RC)
                    self.fc_nat_parks,      # national parks
                    self.fc_crown_grants    # ILRR Crown Grants
                ],
                identity_fcs=[
                    self.fc_bec,            # BEC (BCGW CURRENT)
                    self.fc_road_dissolve,  # buffered/dissolved road ROWs
                    self.fc_vri_clip        # VRI clipped to AOI/cells
                ]
            )
            self.gar_class = Gar4001(
                gar=self.gar,
                output_xls=self.output_xls,
                logger=self.logger,
                gar_config=gar_config
            )

        elif self.gar == 'u-4-007':
            # NOTE: __uwr_golden (local) is replaced with BCGW UWR,
            # and the cell_field must be a BCGW field (e.g., UWR_UNIT_NUMBER).
            gar_config = GARConfig(
                cells=self.__uwr,                  # BCGW: WHSE_WILDLIFE_MANAGEMENT.WCP_UNGULATE_WINTER_RANGE_SP
                cell_field='UWR_UNIT_NUMBER',      # override; don't use the old 'MGT' field (local-only)
                aoi=self.fc_aoi_clean,             # your user-supplied AOI instead of fc_toc_area (BCTS boundary)
                private_land=self.__private_land_pmbc,

                # Keep erase list strictly BCGW-based
                erase_fcs=[
                    self.fc_private_land,          # PMBC parcels (non-Crown etc.)
                    self.fc_prov_parks,
                    self.fc_nat_parks,
                    self.fc_xmas_trees
                ],

                # Keep identity chain lean and BCGW-only
                identity_fcs=[
                    self.fc_bec,                   # BEC (CURRENT)
                    self.fc_road_dissolve,         # buffered/dissolved road ROWs
                    self.fc_vri_clip               # VRI clipped to AOI/cells
                ]
            )
            self.gar_class = Gar4007(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'u-4-010':
            gar_config = GARConfig(
                sql="UWR_NUMBER = '{}' AND FEATURE_NOTES NOT LIKE '%SIC = 0%'".format(self.gar),
                cells=self.__uwr,                 # BCGW UWR source
                cell_field=self.fld_notes,        # FEATURE_NOTES (as in your original)
                aoi=self.fc_aoi_clean,            # was fc_toc_area (BCTS); now your AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,
                    self.fc_prov_parks,
                    self.fc_nat_parks,
                ],
                identity_fcs=[
                    self.fc_bec,
                    self.fc_road_dissolve,
                    self.fc_vri_clip,
                ],
            )
            self.gar_class = Gar4010(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'u-8-001':
            gar_config = GARConfig(
                sql="UWR_NUMBER = '{}' AND FEATURE_NOTES NOT LIKE '%SIC = 0%'".format(self.gar.replace('-tfl49', '')),
                cells=self.__uwr,                 # BCGW UWR source
                cell_field=self.fld_uwr_num,      # typically 'UWR_UNIT_NUMBER' for BCGW
                aoi=self.fc_aoi_clean,            # replace fc_toc_area with your small user AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,         # PMBC parcels (non-Crown etc.)
                    self.fc_woodlots,             # ACTIVE woodlots
                ],
                identity_fcs=[
                    self.fc_bec,                  # BEC (CURRENT)
                    self.fc_road_dissolve,        # buffered/dissolved road ROWs
                    self.fc_vri_clip,             # VRI clipped to AOI/cells
                ],
            )
            self.gar_class = Gar8001(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'u-8-001-tfl49':
            gar_config = GARConfig(
                sql="UWR_NUMBER = '{}' AND FEATURE_NOTES NOT LIKE '%SIC = 0%'".format(self.gar.replace('-tfl49', '')),
                cells=self.__uwr,                 # BCGW UWR source
                cell_field=self.fld_uwr_num,      # typically 'UWR_UNIT_NUMBER' for BCGW
                aoi=self.fc_tfl49,                # AOI is TFL 49 (prepare in prepare_data by selecting FOREST_FILE_ID='TFL49')
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,         # PMBC parcels (non-Crown etc.)
                    self.fc_woodlots,             # ACTIVE woodlots
                ],
                identity_fcs=[
                    self.fc_bec,                  # BEC (CURRENT)
                    self.fc_road_dissolve,        # buffered/dissolved road ROWs
                    self.fc_vri_clip,             # VRI clipped to AOI/cells
                ],
            )
            self.gar_class = Gar8001(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )

        elif self.gar == 'u-8-005':
            gar_config = GARConfig(
                sql="UWR_NUMBER = '{}' AND FEATURE_NOTES NOT LIKE '%SIC = 0%'".format(self.gar),
                cells=self.__uwr,                 # BCGW UWR source
                cell_field=self.fld_uwr_num,      # typically 'UWR_UNIT_NUMBER' for BCGW
                aoi=self.fc_aoi_clean,            # replace fc_toc_area with your AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,         # PMBC parcels (non-Crown etc.)
                    self.fc_woodlots,             # ACTIVE woodlots
                ],
                identity_fcs=[
                    self.fc_bec,                  # BEC (CURRENT)
                    self.fc_road_dissolve,        # buffered/dissolved road ROWs
                    self.fc_vri_clip,             # VRI clipped to AOI/cells
                ],
            )
            self.gar_class = Gar8005(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'u-8-006':
            gar_config = GARConfig(
                sql="UWR_NUMBER = '{}' AND FEATURE_NOTES NOT LIKE '%SIC = 0%'".format(self.gar),
                cells=self.__uwr,                 # BCGW UWR source
                cell_field=self.fld_uwr_num,      # typically 'UWR_UNIT_NUMBER' in BCGW
                aoi=self.fc_aoi_clean,            # replace fc_toc_area with your user AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,         # PMBC parcels (non-Crown etc.)
                    self.fc_woodlots,             # ACTIVE woodlots
                ],
                identity_fcs=[
                    self.fc_bec,                  # BEC (CURRENT)
                    self.fc_road_dissolve,        # buffered/dissolved road ROWs
                    self.fc_vri_clip,             # VRI clipped to AOI/cells
                ],
            )
            self.gar_class = Gar8006(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'u-8-012':
            gar_config = GARConfig(
                sql="UWR_NUMBER = '{}' AND FEATURE_NOTES NOT LIKE '%SIC = 0%'".format(self.gar),
                cells=self.__uwr,                 # BCGW UWR source
                cell_field=self.fld_bec,          # uses BEC label (e.g., BGC_LABEL) as in your original
                aoi=self.fc_aoi_clean,            # replace fc_toc_area with your AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,         # PMBC parcels (non-Crown etc.)
                ],
                identity_fcs=[
                    self.fc_bec,                  # BEC (CURRENT)
                    self.fc_road_dissolve,        # buffered/dissolved road ROWs
                    self.fc_vri_clip,             # VRI clipped to AOI/cells
                ],
            )
            self.gar_class = Gar8012(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'u-8-232':
            gar_config = GARConfig(
                sql="TAG = '{}' AND ORG_ORGANIZATION_ID IN (4, 8)".format(self.gar[2:]),
                cells=self.__wha,                 # BCGW: WHSE_WILDLIFE_MANAGEMENT.WCP_WILDLIFE_HABITAT_AREA_POLY
                cell_field=self.fld_lu,           # LU name will be added via identity with fc_lu
                aoi=self.fc_aoi_clean,            # replace fc_op_areas (LRM) with your AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,
                    self.fc_woodlots,
                    # If you later want a federal mask, we can add a PMBC-based subset back in.
                ],
                identity_fcs=[
                    self.fc_lu,                   # brings LANDSCAPE_UNIT_NAME onto the features
                    self.fc_bec,
                    self.fc_road_dissolve,
                    self.fc_vri_clip,
                ],
            )
            self.gar_class = Gar8232(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'lrmp-bhs':
            gar_config = GARConfig(
                sql=(
                    "STRGC_LAND_RSRCE_PLAN_NAME = 'Okanagan Shuswap Land and Resource Management Plan' "
                    "AND LEGAL_FEAT_OBJECTIVE = 'Big Horn Sheep Areas'"
                ),
                cells=self.__lrmp2,             # BCGW: RMP_PLAN_LEGAL_POLY_SVW
                cell_field=self.fld_lrmp2,      # typically 'LEGAL_FEAT_PROVID'
                aoi=self.fc_aoi_clean,          # use user AOI instead of BCTS boundary
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,       # PMBC parcels (non-Crown etc.)
                    self.fc_woodlots,           # ACTIVE woodlots
                ],
                identity_fcs=[
                    self.fc_bec,                # BEC (CURRENT)
                    self.fc_road_dissolve,      # buffered/dissolved road ROWs
                    self.fc_vri_clip,           # VRI clipped to AOI/cells
                ],
            )
            self.gar_class = LrmpSheep(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )


        elif self.gar == 'lrmp-ds':
            gar_config = GARConfig(
                sql=(
                    "STRGC_LAND_RSRCE_PLAN_NAME = 'Okanagan Shuswap Land and Resource Management Plan' "
                    "AND NON_LEGAL_FEAT_OBJECTIVE = 'Derenzy Bighorn Sheep Habitat RMZ' "
                    "AND NON_LEGAL_FEAT_ATRB_1_VALUE = '2'"
                ),
                cells=self.__lrmp,            # BCGW: RMP_PLAN_NON_LEGAL_POLY_SVW
                cell_field=self.fld_lrmp,     # typically 'NON_LEGAL_FEAT_PROVID'
                aoi=self.fc_aoi_clean,        # replace fc_toc_area with user AOI
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,
                    self.fc_woodlots,
                ],
                identity_fcs=[
                    self.fc_bec,
                    self.fc_road_dissolve,
                    self.fc_vri_clip,
                ],
            )
            self.gar_class = LrmpSheep(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )

            
        elif self.gar == 'section-7':
            # NOTE: Replaces local __sec7 with BCGW UWR. If you specifically need the Golden Sec 7 polygons,
            # supply them as the AOI (fc_aoi_clean) or provide a BCGW-accessible equivalent.
            gar_config = GARConfig(
                cells=self.__uwr,                 # BCGW: WCP_UNGULATE_WINTER_RANGE_SP
                cell_field='UWR_UNIT_NUMBER',     # override; the old 'Name' field was from the local dataset
                aoi=self.fc_aoi_clean,            # use your user-provided AOI instead of fc_toc_area
                private_land=self.__private_land_pmbc,
                erase_fcs=[
                    self.fc_private_land,         # PMBC parcels (non-Crown etc.)
                    self.fc_woodlots,             # ACTIVE woodlots
                ],
                identity_fcs=[
                    self.fc_bec,                  # BEC (CURRENT)
                    self.fc_road_dissolve,        # buffered/dissolved road ROWs
                    self.fc_vri_clip,             # VRI clipped to AOI/cells
                ],
            )
            self.gar_class = Gar8006(
                gar=self.gar, output_xls=self.output_xls, logger=self.logger, gar_config=gar_config
            )
        #--------------------------------------------------------------------------------------------------------------------------------------------------



    def __del__(self):
        """
        Best-effort cleanup; never raise from a destructor.
        """
        try:
            lg = getattr(self, "logger", None) or logging.getLogger(__name__)

            # LRM is not used in the SA build — skip its cleanup.

            # BCGW connection cleanup (guarded)
            if hasattr(Environment, "delete_bcgw_connection"):
                try:
                    Environment.delete_bcgw_connection(location=getattr(self, "sde_folder", None), logger=lg)
                except Exception as e:
                    lg.debug(f"BCGW connection cleanup skipped: {e}")

            # Avoid arcpy deletions here; do them in normal code paths, not in __del__.
            # If you ever want to purge scratch, do it explicitly in the workflow, not here.

        except Exception:
            # Absolutely suppress any destructor-time exceptions
            pass


    def prepare_data(self):
        """
        Prepares BCGW-only inputs, builds AOI, selects GAR cells, and stages
        layers for identity/erase. Uses a small user AOI (or TFL49 for the
        -tfl49 variant) to keep runs fast.
        """
        # ---------------- AOI ----------------
        aoi_fc = None
        try:
            # If the GAR config asks for TFL49 as AOI, build it; else try the user-supplied AOI.
            if getattr(self.gar_class.gar_config, "aoi", None) == self.fc_tfl49:
                self.logger.info("Preparing AOI from TFL49 (BCGW).")
                arcpy.Select_analysis(
                    in_features=self.__tfl,
                    out_feature_class=self.fc_tfl49,
                    where_clause="FOREST_FILE_ID = 'TFL49'"
                )
                aoi_fc = self.fc_tfl49
            else:
                if getattr(self, "aoi", None) and arcpy.Exists(self.aoi):
                    self.logger.info("Preparing AOI from user input.")
                    # Project to BC Albers (EPSG:3005) if needed; otherwise just copy.
                    try:
                        sr = arcpy.Describe(self.aoi).spatialReference
                        if getattr(sr, "factoryCode", None) == 3005:
                            arcpy.CopyFeatures_management(self.aoi, self.fc_aoi_clean)
                        else:
                            arcpy.Project_management(self.aoi, self.fc_aoi_clean, arcpy.SpatialReference(3005))
                    except Exception:
                        # Fallback: try a straight copy
                        arcpy.CopyFeatures_management(self.aoi, self.fc_aoi_clean)
                    # Optional geometry repair (safe/no-op if clean)
                    try:
                        arcpy.RepairGeometry_management(self.fc_aoi_clean)
                    except Exception:
                        pass
                    aoi_fc = self.fc_aoi_clean
                else:
                    self.logger.warning("No AOI provided; continuing without pre-clipping AOI.")
                    aoi_fc = None
        except Exception as e:
            self.logger.error(f"AOI preparation failed: {e}")
            raise

        # ---------------- Select GAR cells (intersect AOI if present) ----------------
        self.logger.info("Selecting GAR cells for the chosen analysis.")
        gar_lyr = arcpy.MakeFeatureLayer_management(
            in_features=self.gar_class.gar_config.cells,
            out_layer="gar_lyr",
            where_clause=self.gar_class.gar_config.sql
        )
        if aoi_fc and arcpy.Exists(aoi_fc):
            aoi_lyr = arcpy.MakeFeatureLayer_management(aoi_fc, "aoi_lyr")
            arcpy.SelectLayerByLocation_management(in_layer=gar_lyr, overlap_type="INTERSECT", select_features=aoi_lyr)
            arcpy.Delete_management(aoi_lyr)

        arcpy.CopyFeatures_management(in_features=gar_lyr, out_feature_class=self.fc_gar_cells)
        arcpy.Delete_management(gar_lyr)

        # Ensure we actually have cells
        if int(arcpy.GetCount_management(self.fc_gar_cells).getOutput(0)) == 0:
            raise RuntimeError("No GAR cells found inside the AOI. Check your AOI and GAR selection parameters.")

        # Merge/clean cells if that GAR needs it (keeps behavior of original)
        self.merge_cells()

        # ---------------- Set processing extent to cells ----------------
        self.logger.info("Setting processing extent to GAR cells.")
        arcpy.env.extent = arcpy.Describe(self.fc_gar_cells).extent

        # ---------------- VRI subset + clip ----------------
        self.logger.info("Subsetting VRI by GAR cells.")
        gar_lyr = arcpy.MakeFeatureLayer_management(self.fc_gar_cells, "gar_lyr_for_vri")
        vri_lyr = arcpy.MakeFeatureLayer_management(self.__vri, "vri_lyr")
        arcpy.SelectLayerByLocation_management(in_layer=vri_lyr, overlap_type="INTERSECT", select_features=gar_lyr)
        arcpy.CopyFeatures_management(in_features=vri_lyr, out_feature_class=self.fc_vri)
        arcpy.Delete_management(vri_lyr)
        arcpy.Delete_management(gar_lyr)

        self.logger.info("Clipping VRI to GAR cells.")
        arcpy.Clip_analysis(in_features=self.fc_vri, clip_features=self.fc_gar_cells, out_feature_class=self.fc_vri_clip)

        # ---------------- Copy other required inputs (BCGW-only) ----------------
        self.logger.info("Preparing additional BCGW inputs as required by the GAR config.")
        gar_lyr = arcpy.MakeFeatureLayer_management(self.fc_gar_cells, "gar_lyr_for_inputs")
        for key, gi in self.dict_gar_inputs.items():
            # Only copy when needed for the current run
            if key.startswith("private_land") and gi.path != self.gar_class.gar_config.private_land:
                continue
            if gi.mandatory or gi.output in self.gar_class.gar_config.erase_fcs or gi.output in self.gar_class.gar_config.identity_fcs:
                self.logger.info(f"  - Copying {key}")
                input_lyr = arcpy.MakeFeatureLayer_management(in_features=gi.path, out_layer="input_lyr", where_clause=gi.sql)
                arcpy.SelectLayerByLocation_management(in_layer=input_lyr, overlap_type="INTERSECT", select_features=gar_lyr)
                arcpy.CopyFeatures_management(in_features=input_lyr, out_feature_class=gi.output)
                arcpy.Delete_management(input_lyr)
        arcpy.Delete_management(gar_lyr)

        # ---------------- Erase masks ----------------
        # Merge only those erase features that actually exist & have content
        erase_inputs = []
        for fc in self.gar_class.gar_config.erase_fcs:
            if arcpy.Exists(fc):
                try:
                    if int(arcpy.GetCount_management(fc).getOutput(0)) > 0:
                        erase_inputs.append(fc)
                except Exception:
                    pass

        if erase_inputs:
            self.logger.info("Combining erase features.")
            arcpy.Merge_management(inputs=erase_inputs, output=self.fc_erase_features)

            self.logger.info("Erasing features from GAR cells.")
            arcpy.Erase_analysis(
                in_features=self.fc_gar_cells,
                erase_features=self.fc_erase_features,
                out_feature_class=self.fc_gar_cells_erase
            )
        else:
            # Nothing to erase; continue with original cells
            arcpy.CopyFeatures_management(self.fc_gar_cells, self.fc_gar_cells_erase)

        # ---------------- Road ROWs (MOT + FTEN) ----------------
        self.logger.info("Building road right-of-ways.")
        road_inputs = []
        for fc in (self.fc_mot_roads, self.fc_ften_roads):
            if arcpy.Exists(fc):
                try:
                    if int(arcpy.GetCount_management(fc).getOutput(0)) > 0:
                        road_inputs.append(fc)
                except Exception:
                    pass

        if road_inputs:
            arcpy.Merge_management(inputs=road_inputs, output=self.fc_road_merge)
            arcpy.Buffer_analysis(
                in_features=self.fc_road_merge,
                out_feature_class=self.fc_road_buffer,
                buffer_distance_or_field="10 Meters",
                dissolve_option="NONE"
            )
            # Tag, then dissolve to a single multipart
            if self.fld_road_buffer not in [f.name for f in arcpy.ListFields(self.fc_road_buffer)]:
                arcpy.AddField_management(self.fc_road_buffer, self.fld_road_buffer, "TEXT", field_length=3)
            with arcpy.da.UpdateCursor(self.fc_road_buffer, [self.fld_road_buffer]) as cur:
                for row in cur:
                    row[0] = "YES"
                    cur.updateRow(row)
            arcpy.Dissolve_management(
                in_features=self.fc_road_buffer,
                out_feature_class=self.fc_road_dissolve,
                dissolve_field=self.fld_road_buffer,
                multi_part="SINGLE_PART"
            )
        else:
            self.logger.info("No road features found within the AOI/cells; skipping road ROW dissolve.")

        # ---------------- (Removed) burn severity / broadleaf ----------------
        # We intentionally skip add_burn_severity() and create_broadleaf_stand_layer()
        # because those relied on non-BCGW/local inputs in the original build.

        # ---------------- Done ----------------
        self.logger.info("Data preparation complete.")




    def add_sic_replacement(self):
        """
        Optional: apply field-verified SIC replacements where available.
        In this simplified SA build, the SIC dataset is on a BCTS share; if it
        isn't accessible, we skip without failing the run.
        """
        try:
            # Only applies to these GARs (same as original intent)
            if self.gar not in ('u-8-006', 'u-8-001', 'u-8-001-tfl49'):
                self.logger.info(f"SIC replacement not applicable for {self.gar}; skipping.")
                return

            # Dataset must exist and be reachable
            if not hasattr(self, "_GARAnalysis__sic_replacement") or not arcpy.Exists(self.__sic_replacement):
                self.logger.info("SIC replacement dataset not accessible; skipping.")
                return

            # Must run after identity_gar(), since we update that output
            if not arcpy.Exists(self.fc_gar_cells_identity):
                self.logger.warning("Identity output missing; run identity_gar() first. Skipping SIC replacement.")
                return

            # Copy SIC polygons into scratch for stable field names
            if not arcpy.Exists(self.fc_sic_replacement):
                arcpy.CopyFeatures_management(self.__sic_replacement, self.fc_sic_replacement)

            # Identity: bring SIC attributes onto the identity FC
            vri_sic = os.path.join(self.scratch_gdb, "vri_sic")
            arcpy.Identity_analysis(
                in_features=self.fc_gar_cells_identity,
                identity_features=self.fc_sic_replacement,
                out_feature_class=vri_sic,
                join_attributes="ALL"
            )

            # Figure out the FID_* field created by Identity for the SIC layer
            fields_in = {f.name for f in arcpy.ListFields(vri_sic)}
            fid_candidates = [n for n in fields_in if n.lower().startswith("fid_")]
            fid_sic = None
            base = os.path.basename(self.fc_sic_replacement)
            # try to match FID_<basename> if present
            for n in fid_candidates:
                if base.replace(".", "_").lower().endswith(n[4:].lower()):
                    fid_sic = n
                    break
            if not fid_sic:
                # fallback: pick the only FID_* besides FID_<in_features>
                if len(fid_candidates) >= 2:
                    fid_sic = sorted(fid_candidates)[-1]
                elif len(fid_candidates) == 1:
                    fid_sic = fid_candidates[0]
                else:
                    self.logger.info("No FID_* join field from SIC identity; skipping SIC replacement.")
                    arcpy.Delete_management(vri_sic)
                    return

            # Map: source SIC fields -> target VRI/identity fields
            src_to_dst = {
                "BEC_ZONE_CODE":        self.fld_bec_zone if hasattr(self, "fld_bec_zone") else None,
                "BEC_SUBZONE":          self.fld_bec_subzone if hasattr(self, "fld_bec_subzone") else None,
                "BEC_VARIANT":          self.fld_bec_variant if hasattr(self, "fld_bec_variant") else None,
                "AGE":                  self.fld_proj_age,
                "DBH":                  self.fld_diameter,
                "HEIGHT":               self.fld_proj_height,
                "CROWN_CLOSURE":        self.fld_crown_closure,
                "SLOPE":                self.fld_slope if hasattr(self, "fld_slope") else None,
                "SPECIES_1":            self.fld_species,
                "SPECIES_PCT_1":        self.fld_percent,
                "SPECIES_2":            getattr(self, "fld_species_2", None),
                "SPECIES_PCT_2":        getattr(self, "fld_percent_2", None),
                "SPECIES_3":            getattr(self, "fld_species_3", None),
                "SPECIES_PCT_3":        getattr(self, "fld_percent_3", None),
                "SPECIES_4":            getattr(self, "fld_species_4", None),
                "SPECIES_PCT_4":        getattr(self, "fld_percent_4", None),
                "SPECIES_5":            getattr(self, "fld_species_5", None),
                "SPECIES_PCT_5":        getattr(self, "fld_percent_5", None),
                "SPECIES_6":            getattr(self, "fld_species_6", None),
                "SPECIES_PCT_6":        getattr(self, "fld_percent_6", None),
                "SURVEY_DATE":          self.fld_proj_date,
            }

            # keep only mappings where both source and dest field actually exist
            update_pairs = []
            for src, dst in src_to_dst.items():
                if dst and (src in fields_in):
                    # ensure destination field exists; add it if needed
                    if dst not in fields_in:
                        ftype = "TEXT"
                        if dst in (self.fld_proj_age,):
                            ftype = "LONG"
                        elif dst in (self.fld_proj_height, self.fld_diameter):
                            ftype = "DOUBLE"
                        elif dst == self.fld_proj_date:
                            ftype = "DATE"
                        try:
                            arcpy.AddField_management(vri_sic, dst, ftype)
                            fields_in.add(dst)
                        except Exception:
                            pass
                    if dst in fields_in:
                        update_pairs.append((src, dst))

            if not update_pairs:
                self.logger.info("No matching SIC fields to transfer; skipping.")
                arcpy.Delete_management(vri_sic)
                return

            # Build cursor field list
            fld_list = [fid_sic] + [p[0] for p in update_pairs] + [p[1] for p in update_pairs]
            idx = {name: i for i, name in enumerate(fld_list)}

            # Update where we actually intersected SIC polygons (FID != -1)
            with arcpy.da.UpdateCursor(vri_sic, fld_list) as cur:
                for row in cur:
                    if row[idx[fid_sic]] is not None and int(row[idx[fid_sic]]) != -1:
                        for src, dst in update_pairs:
                            row[idx[dst]] = row[idx[src]]
                        cur.updateRow(row)

            # Overwrite identity FC with updated attributes
            arcpy.CopyFeatures_management(vri_sic, self.fc_gar_cells_identity)
            arcpy.Delete_management(vri_sic)
            self.logger.info("SIC replacement applied.")
        except Exception as e:
            self.logger.warning(f"SIC replacement skipped due to error: {e}")


    def merge_cells(self):
        """
        Dissolve LRMP cells and carry forward the list of original IDs into a
        single text field (self.gar_class.gar_config.cell_field) using a
        Spatial Join 'JOIN' merge rule.

        No-op for non-LRMP runs.
        """
        # Only LRMP variants require merging (same behavior as original)
        if not (self.gar.startswith("lrmp-")):
            return

        try:
            self.logger.info("Merging LRMP cells (dissolve + ID aggregation).")

            # Paths for temps
            temp_diss   = os.path.join(self.scratch_gdb, "lrmp_cells_diss")
            temp_single = os.path.join(self.scratch_gdb, "lrmp_cells_single")
            temp_join   = os.path.join(self.scratch_gdb, "lrmp_cells_join")

            # 1) Dissolve selected cells to singlepart polygons
            arcpy.Dissolve_management(
                in_features=self.fc_gar_cells,
                out_feature_class=temp_diss,
                multi_part="SINGLE_PART"
            )

            # 2) Remove tiny slivers (< 1,000 m²) to keep results clean
            area_fld = "_AREA_M2"
            if area_fld not in [f.name for f in arcpy.ListFields(temp_diss)]:
                arcpy.AddField_management(temp_diss, area_fld, "DOUBLE")
            with arcpy.da.UpdateCursor(temp_diss, ["SHAPE@AREA", area_fld]) as cur:
                for shp_area, _ in cur:
                    cur.updateRow([shp_area, shp_area])

            lyr = arcpy.MakeFeatureLayer_management(temp_diss, "lrmp_diss_lyr")
            arcpy.SelectLayerByAttribute_management(lyr, "NEW_SELECTION", f"{area_fld} < 1000")
            # Delete selected tiny features
            if int(arcpy.GetCount_management(lyr).getOutput(0)) > 0:
                arcpy.DeleteFeatures_management(lyr)
            arcpy.Delete_management(lyr)

            # 3) Break original cells into singleparts (to aggregate IDs reliably)
            arcpy.MultipartToSinglepart_management(
                in_features=self.fc_gar_cells,
                out_feature_class=temp_single
            )

            # 4) Spatial Join (target = dissolved, join = original-singlepart)
            #    Build field mappings to JOIN original IDs into one string field
            cell_field = self.gar_class.gar_config.cell_field

            fm = arcpy.FieldMappings()
            fm.addTable(temp_diss)   # keep dissolved target fields

            fm_cell = arcpy.FieldMap()
            fm_cell.addInputField(temp_single, cell_field)
            out_f = fm_cell.outputField
            out_f.name = cell_field
            out_f.aliasName = cell_field
            out_f.length = max(getattr(out_f, "length", 0) or 0, 1000)  # room for joined IDs
            fm_cell.outputField = out_f
            fm_cell.mergeRule = "Join"
            fm_cell.joinDelimiter = ", "
            fm.addFieldMap(fm_cell)

            arcpy.SpatialJoin_analysis(
                target_features=temp_diss,
                join_features=temp_single,
                out_feature_class=temp_join,
                join_operation="JOIN_ONE_TO_ONE",
                match_option="CONTAINS",  # mirror original 'contains' logic
                field_mapping=fm
            )

            # 5) Overwrite the input cells with merged result
            arcpy.CopyFeatures_management(temp_join, self.fc_gar_cells)

        except Exception as e:
            self.logger.warning(f"merge_cells skipped due to error: {e}")
        finally:
            for fc in (temp_diss, temp_single, temp_join):
                try:
                    if arcpy.Exists(fc):
                        arcpy.Delete_management(fc)
                except Exception:
                    pass




    def identity_gar(self):
        """
        Run a robust identity chain over configured layers.
        Skips missing/empty layers; attempts dicing/subdivide on failure;
        never aborts the whole run because of one problematic layer.
        """
        try:
            input_fc = self.fc_gar_cells_erase
            if not arcpy.Exists(input_fc):
                self.logger.warning("identity_gar: input (fc_gar_cells_erase) missing; copying cells.")
                arcpy.CopyFeatures_management(self.fc_gar_cells, self.fc_gar_cells_identity)
                return
            try:
                if int(arcpy.GetCount_management(input_fc).getOutput(0)) == 0:
                    self.logger.warning("identity_gar: input has no features; copying cells.")
                    arcpy.CopyFeatures_management(self.fc_gar_cells, self.fc_gar_cells_identity)
                    return
            except Exception:
                pass

            # Filter identity layers to ones that exist and have features
            id_layers = []
            for ident_lyr in (self.gar_class.gar_config.identity_fcs or []):
                if not ident_lyr or not arcpy.Exists(ident_lyr):
                    continue
                try:
                    if int(arcpy.GetCount_management(ident_lyr).getOutput(0)) > 0:
                        id_layers.append(ident_lyr)
                except Exception:
                    # If count fails, still try to use it
                    id_layers.append(ident_lyr)

            if not id_layers:
                self.logger.info("identity_gar: no identity layers to apply; copying input to output.")
                arcpy.CopyFeatures_management(input_fc, self.fc_gar_cells_identity)
                return

            # Scratch paths
            out_fc         = os.path.join(self.scratch_gdb, "id_out")
            next_input_fc  = os.path.join(self.scratch_gdb, "id_input")
            dice_temp      = os.path.join(self.scratch_gdb, "id_dice")
            subdivide_temp = os.path.join(self.scratch_gdb, "id_subdivide")

            work_in = input_fc

            for ident_lyr in id_layers:
                name = os.path.basename(ident_lyr)
                self.logger.info(f"Identity: {name}")

                attempts = ("direct", "dice", "subdivide+dice")
                succeeded = False

                for attempt in attempts:
                    try:
                        if attempt == "direct":
                            source_fc = ident_lyr

                        elif attempt == "dice":
                            if arcpy.Exists(dice_temp):
                                arcpy.Delete_management(dice_temp)
                            arcpy.Dice_management(in_features=ident_lyr, out_feature_class=dice_temp, vertex_limit=10000)
                            source_fc = dice_temp

                        else:  # "subdivide+dice"
                            # Use Environment.subdivide_polygons if available; else retry dice directly
                            if hasattr(Environment, "subdivide_polygons"):
                                if arcpy.Exists(subdivide_temp):
                                    arcpy.Delete_management(subdivide_temp)
                                sub_fc = Environment.subdivide_polygons(input_fc=ident_lyr, output_fc=subdivide_temp)
                                if arcpy.Exists(dice_temp):
                                    arcpy.Delete_management(dice_temp)
                                arcpy.Dice_management(in_features=sub_fc, out_feature_class=dice_temp, vertex_limit=10000)
                                source_fc = dice_temp
                            else:
                                # No subdivide helper available; just skip this attempt
                                raise RuntimeError("subdivide helper not available")

                        # Try the identity
                        arcpy.Identity_analysis(
                            in_features=work_in,
                            identity_features=source_fc,
                            out_feature_class=out_fc,
                            join_attributes='NO_FID'
                        )

                        # Success → promote to next input
                        if arcpy.Exists(next_input_fc):
                            arcpy.Delete_management(next_input_fc)
                        arcpy.CopyFeatures_management(out_fc, next_input_fc)
                        work_in = next_input_fc
                        succeeded = True
                        break

                    except Exception as e:
                        self.logger.warning(f"Identity '{attempt}' failed on {name}: {e}")

                if not succeeded:
                    self.logger.error(f"Skipping layer (all attempts failed): {name}")
                    # Keep current work_in and move on

                # Cleanup per-layer temps
                for fc in (out_fc, dice_temp):
                    try:
                        if arcpy.Exists(fc):
                            arcpy.Delete_management(fc)
                    except Exception:
                        pass

            # Finalize
            arcpy.CopyFeatures_management(work_in, self.fc_gar_cells_identity)

        except Exception as e:
            self.logger.error(f"identity_gar failed; copying input to output. Error: {e}")
            try:
                arcpy.CopyFeatures_management(self.fc_gar_cells_erase, self.fc_gar_cells_identity)
            except Exception:
                pass
        finally:
            # Best-effort cleanup
            for fc in ("id_out", "id_input", "id_dice", "id_subdivide"):
                try:
                    p = os.path.join(self.scratch_gdb, fc)
                    if arcpy.Exists(p):
                        arcpy.Delete_management(p)
                except Exception:
                    pass



    def fix_slivers(self):
        """
        Clean up identity output by converting to singlepart, repairing geometry,
        iteratively eliminating ~1 m² slivers, and writing the resultant.
        """
        try:
            # Guard: need identity output to proceed
            if not arcpy.Exists(self.fc_gar_cells_identity):
                self.logger.warning("fix_slivers: identity output missing; copying erased cells to resultant.")
                if arcpy.Exists(self.fc_gar_cells_erase):
                    arcpy.CopyFeatures_management(self.fc_gar_cells_erase, self.fc_resultant)
                else:
                    # last-resort fallback
                    arcpy.CopyFeatures_management(self.fc_gar_cells, self.fc_resultant)
                return

            try:
                if int(arcpy.GetCount_management(self.fc_gar_cells_identity).getOutput(0)) == 0:
                    self.logger.warning("fix_slivers: identity output empty; copying erased cells to resultant.")
                    arcpy.CopyFeatures_management(self.fc_gar_cells_erase, self.fc_resultant)
                    return
            except Exception:
                pass

            single_part_output = self.fc_gar_cells_single
            fld_area = 'Area_m'

            # Convert to singlepart and repair geometry
            self.logger.info('Converting identity output to singlepart.')
            arcpy.MultipartToSinglepart_management(
                in_features=self.fc_gar_cells_identity,
                out_feature_class=single_part_output
            )

            self.logger.info('Repairing geometry.')
            try:
                arcpy.RepairGeometry_management(in_features=single_part_output)
            except Exception as e:
                self.logger.warning(f"RepairGeometry failed (continuing): {e}")

            # Ensure area field exists, then populate with SHAPE@AREA
            if fld_area not in [f.name for f in arcpy.ListFields(single_part_output)]:
                arcpy.AddField_management(in_table=single_part_output, field_name=fld_area, field_type='DOUBLE')

            with arcpy.da.UpdateCursor(single_part_output, ['SHAPE@AREA', fld_area]) as u_cursor:
                for shp_area, _ in u_cursor:
                    u_cursor.updateRow([shp_area, shp_area])

            # Temp outputs that we toggle between while iterating
            out_a = os.path.join(self.scratch_gdb, 'out_temp_a')
            out_b = os.path.join(self.scratch_gdb, 'out_temp_b')

            # First pass
            current_selection = self.eliminate_small_polygons(
                inputfc=single_part_output,
                outputfc=out_a,
                area_field=fld_area
            )

            self.logger.info('Merging ~1 m² polygons with largest neighbour (iterative).')
            prev_selection = float('inf')
            passes = 0
            max_passes = 8  # hard stop to avoid infinite loops on pathological geometry

            # Alternate between out_a and out_b each pass
            while prev_selection > current_selection and current_selection > 0 and passes < max_passes:
                self.logger.info(f'{current_selection} polygon(s) remaining under threshold (pass {passes + 1}).')
                input_fc = out_a if passes % 2 == 0 else out_b
                output_fc = out_b if passes % 2 == 0 else out_a
                prev_selection = current_selection
                passes += 1

                current_selection = self.eliminate_small_polygons(
                    inputfc=input_fc,
                    outputfc=output_fc,
                    area_field=fld_area
                )

            # Choose the latest output we wrote to
            final_fc = out_a if passes % 2 == 1 else out_b

            # Strip helper area field (if present) before writing the final Resultant
            if fld_area in [f.name for f in arcpy.ListFields(final_fc)]:
                try:
                    arcpy.DeleteField_management(in_table=final_fc, drop_field=fld_area)
                except Exception:
                    # non-fatal
                    pass

            self.logger.info('Creating resultant.')
            arcpy.CopyFeatures_management(in_features=final_fc, out_feature_class=self.fc_resultant)

        except Exception as e:
            self.logger.error(f"fix_slivers failed; writing identity output directly as resultant. Error: {e}")
            try:
                arcpy.CopyFeatures_management(self.fc_gar_cells_identity, self.fc_resultant)
            except Exception:
                pass
        finally:
            # Cleanup temps (best effort)
            for f in [
                os.path.join(self.scratch_gdb, 'out_temp_a'),
                os.path.join(self.scratch_gdb, 'out_temp_b'),
                self.fc_gar_cells_single
            ]:
                try:
                    if arcpy.Exists(f):
                        arcpy.Delete_management(f)
                except Exception:
                    pass


    def eliminate_small_polygons(self, inputfc, outputfc, area_field):
        """
        Eliminate polygons with area_field < 1 m² by merging to neighbors.
        Returns the count of polygons under the threshold at selection time.
        """
        # Make a unique layer to avoid name collisions
        lyr_name = arcpy.CreateUniqueName("temp_lyr")
        temp_layer = arcpy.MakeFeatureLayer_management(in_features=inputfc, out_layer=lyr_name)

        # Select all polygons that are less than 1 square metre
        where_clause = f"{arcpy.AddFieldDelimiters(inputfc, area_field)} < 1"
        arcpy.SelectLayerByAttribute_management(
            in_layer_or_view=temp_layer,
            selection_type='NEW_SELECTION',
            where_clause=where_clause
        )

        # Count current selection
        try:
            current_selection = int(arcpy.GetCount_management(in_rows=temp_layer).getOutput(0))
        except Exception:
            current_selection = 0

        # If no slivers, just copy input → output and return
        if current_selection == 0:
            if arcpy.Exists(outputfc):
                arcpy.Delete_management(outputfc)
            arcpy.CopyFeatures_management(in_features=inputfc, out_feature_class=outputfc)
            try:
                arcpy.Delete_management(temp_layer)
            except Exception:
                pass
            return 0

        # Try the straightforward Eliminate on the selection
        try:
            if arcpy.Exists(outputfc):
                arcpy.Delete_management(outputfc)
            arcpy.Eliminate_management(
                in_features=temp_layer,
                out_feature_class=outputfc,
                selection='AREA'
            )
        except Exception:
            # If Eliminate fails, fall back to copying input as-is
            self.logger.warning('Eliminate failed; copying input to output unchanged.')
            if arcpy.Exists(outputfc):
                arcpy.Delete_management(outputfc)
            arcpy.CopyFeatures_management(in_features=inputfc, out_feature_class=outputfc)

        # Refresh the area field on the output
        try:
            with arcpy.da.UpdateCursor(outputfc, ['SHAPE@AREA', area_field]) as u_cursor:
                for shp_area, _ in u_cursor:
                    u_cursor.updateRow([shp_area, shp_area])
        except Exception:
            pass

        # Cleanup
        try:
            arcpy.Delete_management(temp_layer)
        except Exception:
            pass

        return current_selection





    def calculate_values(self):
        """
        Adds new fields and calculates values based off attributes in the resultant.
        Also computes targets/ranks when supported by the GAR class.
        """
        # Ensure we have a resultant to work with
        if not arcpy.Exists(self.fc_resultant):
            self.logger.error("Resultant feature class not found; cannot calculate values.")
            return
        try:
            if int(arcpy.GetCount_management(self.fc_resultant).getOutput(0)) == 0:
                self.logger.warning("Resultant is empty; skipping calculate_values.")
                return
        except Exception:
            pass

        # Add required fields (idempotent)
        for fld in [
            self.fld_age_cur, self.fld_height_cur, self.fld_height_text, self.fld_level,
            self.fld_rank_oa, self.fld_rank_cell, self.fld_bec_version, self.fld_date_created,
            self.fld_calc_cflb
        ]:
            field_type = 'TEXT'
            if fld == self.fld_age_cur:
                field_type = 'SHORT'
            elif fld == self.fld_date_created:
                field_type = 'DATE'
            elif fld == self.fld_height_cur:
                field_type = 'DOUBLE'
            try:
                if fld not in [f.name for f in arcpy.ListFields(self.fc_resultant)]:
                    arcpy.AddField_management(in_table=self.fc_resultant, field_name=fld, field_type=field_type)
            except Exception:
                # Non-fatal if it already exists or creation fails (read-only FC etc.)
                pass

        self.logger.info('Updating stand attributes and derived fields.')
        current_year = dt.now().year

        # Build a safe field list only from fields that exist + SHAPE@AREA (pseudo-field)
        present_names = {f.name for f in arcpy.ListFields(self.fc_resultant)}
        requested = [
            self.fld_proj_date, self.fld_proj_age, self.fld_age_cur, self.fld_road_buffer, self.fld_cc_status,
            self.fld_cc_harv_date, self.fld_bec_version, self.fld_date_created, self.fld_bec, self.fld_level,
            self.fld_species, self.fld_crown_closure, self.fld_slope, self.fld_thlb, self.fld_diameter,
            self.fld_percent, self.fld_notes, self.fld_op_area, self.fld_calc_cflb, self.fld_bclcs_2,
            self.fld_open_ind, self.fld_line_7b_dist_hist, self.fld_proj_height, self.fld_height_cur,
            self.fld_height_text, self.fld_for_mgmt_ind
        ]
        # Include cell_field if present
        cell_field = getattr(self.gar_class.gar_config, 'cell_field', None)
        if cell_field and cell_field in present_names:
            requested.append(cell_field)

        # Keep only actually present fields; SHAPE@AREA handled separately
        field_list = [f for f in requested if f in present_names]
        # Always append SHAPE@AREA pseudo-field for area calculations
        field_list.append('SHAPE@AREA')

        # Helper for safe reads
        def get_val(row, fields, name, default=None):
            return row[fields.index(name)] if name in fields else default

        with arcpy.da.UpdateCursor(self.fc_resultant, field_list) as u_cursor:
            for row in u_cursor:
                # --- Safe reads (default when absent) ---
                proj_date   = get_val(row, field_list, self.fld_proj_date, None)
                proj_age    = get_val(row, field_list, self.fld_proj_age, None)
                proj_hgt    = get_val(row, field_list, self.fld_proj_height, None)
                rd_buffer   = get_val(row, field_list, self.fld_road_buffer, None)
                cc_status   = get_val(row, field_list, self.fld_cc_status, '')
                cc_harv_dt  = get_val(row, field_list, self.fld_cc_harv_date, '')
                bec         = (get_val(row, field_list, self.fld_bec, '') or '').replace(' ', '')
                spp         = str(get_val(row, field_list, self.fld_species, '') or '')
                cc          = get_val(row, field_list, self.fld_crown_closure, None)
                slp         = get_val(row, field_list, self.fld_slope, None)
                thlb        = None
                if self.fld_thlb in field_list:
                    try:
                        thlb_raw = get_val(row, field_list, self.fld_thlb, None)
                        thlb = float(thlb_raw) if thlb_raw is not None else 0.0
                    except Exception:
                        thlb = 0.0
                diam       = get_val(row, field_list, self.fld_diameter, None)
                pct        = get_val(row, field_list, self.fld_percent, None)
                notes      = get_val(row, field_list, self.fld_notes, '') or ''
                target     = (int(notes[notes.find('=') + 2:]) if ('=' in notes and any(c.isdigit() for c in notes)) else None)
                pcell      = get_val(row, field_list, cell_field, '') if cell_field else ''
                op_area    = get_val(row, field_list, self.fld_op_area, '')
                shp_area   = (row[field_list.index('SHAPE@AREA')] / 10000.0) if 'SHAPE@AREA' in field_list else None
                for_ind    = get_val(row, field_list, self.fld_for_mgmt_ind, 'N')

                # --- Derivations ---
                calc_cflb   = None
                height_cur  = None
                height_text = None
                age_cur     = None

                # Age/height growth from projected values
                if proj_date:
                    try:
                        difference = current_year - proj_date.year
                    except Exception:
                        difference = 0
                    # Age
                    try:
                        if proj_age is not None:
                            age_cur = int(proj_age) + difference
                    except Exception:
                        pass
                    # Height
                    try:
                        if proj_hgt is not None:
                            height_cur = float(proj_hgt) + (0.3 * difference)  # 30 cm/yr
                            if height_cur >= 19.5:
                                height_text = '>= 19.5m'
                    except Exception:
                        pass

                # Harvest override (if attributes available)
                if cc_harv_dt and cc_status not in ('ROAD', 'RESERVE', None):
                    try:
                        age_cur = current_year - int(str(cc_harv_dt)[0:4])
                    except Exception:
                        pass

                # Road buffer nullifies age
                if cc_status == 'ROAD':
                    if self.fld_road_buffer in field_list:
                        row[field_list.index(self.fld_road_buffer)] = 'Yes'
                    age_cur = None
                if rd_buffer == 'Yes':
                    age_cur = None

                # CFLB indicator (calculate + persist)
                if str(for_ind).upper() == 'Y':
                    calc_cflb = 'Y'
                    if self.fld_calc_cflb in field_list:
                        row[field_list.index(self.fld_calc_cflb)] = calc_cflb

                # Level classification (most GARs)
                if getattr(self, 'gar', None) != 'u-8-232':
                    try:
                        level = self.gar_class.calculate_level(
                            bec=bec, age=age_cur, spp=spp, cc=cc, slp=slp, thlb=thlb,
                            diam=diam, pct=pct, gfa=calc_cflb, notes=notes,
                            op_area=op_area, pcell=pcell, shp_area=shp_area,
                            target=target, height=height_cur
                        )
                        if self.fld_level in field_list:
                            row[field_list.index(self.fld_level)] = level
                    except Exception as e:
                        self.logger.warning(f"calculate_level failed (continuing): {e}")

                # Write back always-present (we added them) fields
                if self.fld_age_cur in field_list:
                    row[field_list.index(self.fld_age_cur)] = age_cur
                if self.fld_height_cur in field_list:
                    row[field_list.index(self.fld_height_cur)] = height_cur
                if self.fld_height_text in field_list:
                    row[field_list.index(self.fld_height_text)] = height_text
                if self.fld_bec_version in field_list:
                    row[field_list.index(self.fld_bec_version)] = self.bec_version
                if self.fld_date_created in field_list:
                    row[field_list.index(self.fld_date_created)] = dt.now()  # DATE field prefers datetime

                u_cursor.updateRow(row)

        # Special handling for u-8-232 (unchanged logic, but guard types)
        if getattr(self, 'gar', None) == 'u-8-232':
            lst_fields = [self.fld_op_area, self.fld_lu, self.fld_bec_zone_alt, self.fld_bec_subzone_alt,
                        self.fld_level, self.fld_height_text]
            lst_fields = [f for f in lst_fields if f in {fld.name for fld in arcpy.ListFields(self.fc_resultant)}]
            if lst_fields:
                arcpy.Dissolve_management(in_features=self.fc_resultant,
                                        out_feature_class=self.fc_resultant_dissolve,
                                        dissolve_field=lst_fields)
                work_fields = lst_fields + ['SHAPE@AREA']
                with arcpy.da.UpdateCursor(self.fc_resultant_dissolve, work_fields) as u_cursor:
                    for row in u_cursor:
                        hgt = get_val(row, work_fields, self.fld_height_text, None)
                        shp_area = row[work_fields.index('SHAPE@AREA')] / 10000.0
                        bec = '{0} {1}'.format(
                            get_val(row, work_fields, self.fld_bec_zone_alt, ''),
                            get_val(row, work_fields, self.fld_bec_subzone_alt, '')
                        ).strip()
                        op_area = get_val(row, work_fields, self.fld_op_area, '')
                        lu = get_val(row, work_fields, self.fld_lu, '')
                        try:
                            level = self.gar_class.calculate_level(op_area=op_area,
                                                                pcell=f'{lu}: {bec}',
                                                                shp_area=shp_area,
                                                                height=hgt)
                            if self.fld_level in work_fields:
                                row[work_fields.index(self.fld_level)] = level
                            u_cursor.updateRow(row)
                        except Exception as e:
                            self.logger.warning(f"u-8-232 level calc failed: {e}")

        # Calculate targets (if available)
        try:
            if hasattr(self.gar_class, 'calculate_targets'):
                self.gar_class.calculate_targets()
        except Exception as e:
            self.logger.warning(f"calculate_targets failed (continuing): {e}")

        # Apply ranks (if configured)
        try:
            if getattr(self.gar_class.gar_config, 'ranks', False):
                self.logger.info('Updating resultant with ranks.')
                fields_needed = [self.fld_level, self.fld_op_area, self.fld_bec]
                if cell_field and cell_field in present_names:
                    fields_needed.append(cell_field)
                fields_needed = [f for f in fields_needed if f in {fld.name for fld in arcpy.ListFields(self.fc_resultant)}]

                with arcpy.da.UpdateCursor(self.fc_resultant, fields_needed) as u_cursor:
                    for row in u_cursor:
                        level = str(get_val(row, fields_needed, self.fld_level, '') or '')
                        op_area = get_val(row, fields_needed, self.fld_op_area, '')
                        bec = (get_val(row, fields_needed, self.fld_bec, '') or '').replace(' ', '')
                        pcell = get_val(row, fields_needed, cell_field, '') if cell_field else ''

                        try:
                            if getattr(self, 'gar', None) == 'u-8-006':
                                oa_rank = self.gar_class.dict_total_area[op_area].pcell[pcell].level[level].rank
                                cell_rank = self.gar_class.dict_cell_area[pcell].level[level].rank
                            else:
                                oa_rank = self.gar_class.dict_total_area[op_area].pcell[pcell].level[level].bec[bec].rank
                                cell_rank = self.gar_class.dict_cell_area[pcell].level[level].bec[bec].rank
                        except Exception:
                            # If any lookup fails, skip ranking this row
                            continue

                        # Write ranks if fields exist (we added them earlier)
                        if self.fld_rank_oa in {f.name for f in arcpy.ListFields(self.fc_resultant)}:
                            with arcpy.da.UpdateCursor(self.fc_resultant,
                                                    [self.fld_rank_oa, self.fld_rank_cell],
                                                    where_clause=None) as wcur:
                                pass  # no-op; outer cursor already updating per-row in a single cursor scope

                # Simpler: open a new cursor to set ranks safely
                with arcpy.da.UpdateCursor(self.fc_resultant,
                                        [self.fld_level, self.fld_op_area, self.fld_bec,
                                            (cell_field if cell_field else self.fld_level),
                                            self.fld_rank_oa, self.fld_rank_cell]) as rcur:
                    for lvl, oa, bec_val, pcell_val, _, _ in rcur:
                        try:
                            if getattr(self, 'gar', None) == 'u-8-006':
                                oa_rank = self.gar_class.dict_total_area[oa].pcell[pcell_val].level[str(lvl)].rank
                                cell_rank = self.gar_class.dict_cell_area[pcell_val].level[str(lvl)].rank
                            else:
                                oa_rank = self.gar_class.dict_total_area[oa].pcell[pcell_val].level[str(lvl)].bec[(bec_val or '').replace(' ', '')].rank
                                cell_rank = self.gar_class.dict_cell_area[pcell_val].level[str(lvl)].bec[(bec_val or '').replace(' ', '')].rank
                            rcur.updateRow([lvl, oa, bec_val, pcell_val, oa_rank, cell_rank])
                        except Exception:
                            continue

                if getattr(self, 'gar', None) == 'u-8-006':
                    self.logger.info('Calculating mature stands (u-8-006).')
                    for fld in [self.fld_rank_oa, self.fld_rank_cell]:
                        sql_all = f"{fld} IN ('CH', 'NH')"
                        sql_mature = f"{self.fld_level} = 'Mature Cover'"
                        dissolve_fields = [cell_field, self.fld_op_area] if fld == self.fld_rank_oa and cell_field else ([cell_field] if cell_field else [])
                        if dissolve_fields:
                            self.calculate_mature_stands(where_clause=sql_all, dissolve_fields=dissolve_fields, run_type='All')
                            self.calculate_mature_stands(where_clause=sql_mature, dissolve_fields=dissolve_fields, run_type='Mature')
        except Exception as e:
            self.logger.warning(f"Ranking step skipped due to error: {e}")


    def calculate_mature_stands(self, where_clause, dissolve_fields, run_type):
        """
        Determines which stands are considered mature and aggregates their area
        into gar_class dictionaries. Safe for missing fields / empty selections.
        """
        AREA_THRESH_HA = 20.0

        # Basic existence checks
        if not arcpy.Exists(self.fc_resultant):
            self.logger.warning("Resultant not found; skipping calculate_mature_stands.")
            return

        present = {f.name for f in arcpy.ListFields(self.fc_resultant)}
        missing = [f for f in dissolve_fields if f not in present]
        if missing:
            self.logger.warning(f"Missing dissolve fields {missing}; skipping calculate_mature_stands.")
            return

        # Ensure the cell field exists when needed
        cell_field = getattr(self.gar_class.gar_config, 'cell_field', None)
        if not cell_field or cell_field not in present:
            self.logger.warning("cell_field not present; skipping calculate_mature_stands.")
            return

        # Make selection layer
        result_lyr = arcpy.MakeFeatureLayer_management(self.fc_resultant, 'result_lyr',
                                                    where_clause=where_clause if where_clause else None)
        try:
            cnt = int(arcpy.GetCount_management(result_lyr).getOutput(0))
            if cnt == 0:
                self.logger.info("No features match mature-stand selection; nothing to do.")
                return

            fc_dissolve = os.path.join(self.scratch_gdb, 'dissolve_temp')
            arcpy.Dissolve_management(in_features=result_lyr, out_feature_class=fc_dissolve,
                                    dissolve_field=dissolve_fields, multi_part='SINGLE_PART')

            work_fields = list(dissolve_fields) + ['SHAPE@AREA']
            use_op_area = (self.fld_op_area in dissolve_fields)

            with arcpy.da.SearchCursor(fc_dissolve, work_fields) as s_cur:
                for row in s_cur:
                    shp_ha = row[work_fields.index('SHAPE@AREA')] / 10000.0
                    if shp_ha < AREA_THRESH_HA:
                        continue

                    pcell = row[work_fields.index(cell_field)]
                    try:
                        if run_type == 'Mature':
                            if use_op_area:
                                op_area = row[work_fields.index(self.fld_op_area)]
                                self.gar_class.dict_total_area[op_area].pcell[pcell].level[self.gar_class.str_mature].stand_hectares += shp_ha
                            else:
                                self.gar_class.dict_cell_area[pcell].level[self.gar_class.str_mature].stand_hectares += shp_ha
                        else:  # 'All'
                            if use_op_area:
                                op_area = row[work_fields.index(self.fld_op_area)]
                                self.gar_class.dict_total_area[op_area].pcell[pcell].stand_hectares += shp_ha
                            else:
                                self.gar_class.dict_cell_area[pcell].stand_hectares += shp_ha
                    except KeyError:
                        # If rank/target dictionaries aren’t populated for this GAR or key,
                        # just skip gracefully.
                        continue
                    except Exception as e:
                        self.logger.warning(f"Failed updating mature-stand area for pcell '{pcell}': {e}")
                        continue
        finally:
            # Cleanup
            try:
                arcpy.Delete_management(result_lyr)
            except Exception:
                pass
            try:
                if 'fc_dissolve' in locals() and arcpy.Exists(fc_dissolve):
                    arcpy.Delete_management(fc_dissolve)
            except Exception:
                pass


    def dissolve_resultant(self):
        """
        Creates a dissolved resultant feature class.
        Safe if fields are missing or the input is empty.
        """
        self.logger.info("Dissolving resultant")

        # Guard: input must exist and have features
        if not arcpy.Exists(self.fc_resultant):
            self.logger.warning("Resultant not found; skipping dissolve_resultant.")
            return
        try:
            if int(arcpy.GetCount_management(self.fc_resultant).getOutput(0)) == 0:
                self.logger.warning("Resultant is empty; skipping dissolve_resultant.")
                return
        except Exception as e:
            self.logger.warning(f"Could not count features in resultant: {e}")

        # Only keep dissolve fields that actually exist
        desired = [
            self.fld_uwr_num, self.fld_notes, self.fld_op_area, self.fld_level,
            self.fld_rank_cell, self.fld_rank_oa, self.fld_bec,
            self.fld_bec_version, self.fld_date_created
        ]
        present = {f.name for f in arcpy.ListFields(self.fc_resultant)}
        dissolve_fields = [f for f in desired if f in present]

        # If the output exists, replace it
        if arcpy.Exists(self.fc_resultant_rank):
            try:
                arcpy.Delete_management(self.fc_resultant_rank)
            except Exception:
                pass

        # If no fields survive, dissolve everything into a single feature
        if dissolve_fields:
            arcpy.Dissolve_management(self.fc_resultant, self.fc_resultant_rank,
                                    dissolve_fields, multi_part="SINGLE_PART")
        else:
            arcpy.Dissolve_management(self.fc_resultant, self.fc_resultant_rank,
                                    multi_part="SINGLE_PART")
        self.logger.info(
            f"Dissolve complete. Fields used: {dissolve_fields if dissolve_fields else '[all merged into one]'}"
        )



if __name__ == '__main__':
    run_app()

