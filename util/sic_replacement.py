import arcpy
import os
import sys
import logging

from argparse import ArgumentParser
from datetime import datetime as dt

sys.path.insert(1, r'\\spatialfiles2.bcgov\work\FOR\RSI\TOC\Projects\ESRI_Scripts\Python_Repository')
from environment import Environment


def run_app():
    in_poly, zone, sub, var, age, dbh, hgt, cc, slp, sp1, per1, sp2, per2, sp3, per3, \
    sp4, per4, sp5, per5, sp6, per6, survey_dt, logger = get_input_parameters()
    sic = SICReplacement(in_poly=in_poly, zone=zone, sub=sub, var=var, age=age, dbh=dbh, hgt=hgt, cc=cc, slp=slp,
                         sp1=sp1, per1=per1, sp2=sp2, per2=per2, sp3=sp3, per3=per3, sp4=sp4, per4=per4,
                         sp5=sp5, per5=per5, sp6=sp6, per6=per6, survey_dt=survey_dt, logger=logger)
    sic.replace_sic()
    del sic


def get_input_parameters():
    try:
        parser = ArgumentParser(
            description='This script takes polygons and attributes as input and adds them to a SIC replacement areas '
                        'feature class for use in GAR analysis')
        parser.add_argument('in_poly', type=str, help='Input selected polygon')
        parser.add_argument('zone', type=str, help='BEC Zone Code')
        parser.add_argument('sub', type=str, help='BEC Subzone Code')
        parser.add_argument('var', type=str, help='BEC Variant')
        parser.add_argument('age', type=str, help='Age')
        parser.add_argument('dbh', type=str, help='Diameter')
        parser.add_argument('hgt', type=str, help='Height')
        parser.add_argument('cc', type=str, help='Crown Closure')
        parser.add_argument('slp', type=str, help='Slope')
        parser.add_argument('sp1', type=str, help='Timber type species 1')
        parser.add_argument('per1', type=str, help='Timber type species percent 1')
        parser.add_argument('sp2', type=str, help='Timber type species 2')
        parser.add_argument('per2', type=str, help='Timber type species percent 2')
        parser.add_argument('sp3', type=str, help='Timber type species 3')
        parser.add_argument('per3', type=str, help='Timber type species percent 3')
        parser.add_argument('sp4', type=str, help='Timber type species 4')
        parser.add_argument('per4', type=str, help='Timber type species percent 4')
        parser.add_argument('sp5', type=str, help='Timber type species 5')
        parser.add_argument('per5', type=str, help='Timber type species percent 5')
        parser.add_argument('sp6', type=str, help='Timber type species 6')
        parser.add_argument('per6', type=str, help='Timber type species percent 6')
        parser.add_argument('dt', type=str, help='Survey Date')

        parser.add_argument('--log_level', default='INFO', choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                            help='Log level')
        parser.add_argument('--log_dir', help='Path to log directory')

        args = parser.parse_args()

        logger = Environment.setup_logger(args)

        python_script = sys.argv[0]
        python_dir = os.path.dirname(python_script)

        return args.in_poly, args.zone, args.sub, args.var, args.age, args.dbh, args.hgt, args.cc, args.slp, args.sp1, \
            args.per1, args.sp2, args.per2, args.sp3, args.per3, args.sp4, args.per4, args.sp5, args.per5, \
            args.sp6, args.per6, args.dt, logger

    except Exception as e:
        error_string = 'Unexpected exception. Program terminating: {}'.format(e.message)
        logging.error(msg=error_string)
        arcpy.AddError(message=error_string)
        raise Exception('Errors exist')


class SICReplacement:
    def __init__(self, in_poly, zone, sub, var, age, dbh, hgt, cc, slp, sp1, per1, sp2, per2, sp3, per3, sp4, per4,
                 sp5, per5, sp6, per6, survey_dt, logger):
        arcpy.env.overwriteOutput = True
        self.scratch_gdb = 'in_memory'
        self.in_poly = in_poly
        self.zone = str(zone).upper()
        self.sub = str(sub).lower()
        self.var = str(var)
        self.age = int(age)
        self.dbh = float(dbh)
        self.hgt = float(hgt)
        self.cc = int(cc)
        self.slp = int(slp)
        self.sp1 = str(sp1).upper()
        self.per1 = int(per1)
        self.sp2 = str(sp2).upper() if str(sp2) != '#' else None
        self.per2 = int(per2) if str(per2) != '#' else None
        self.sp3 = str(sp3).upper() if str(sp3) != '#' else None
        self.per3 = int(per3) if str(per3) != '#' else None
        self.sp4 = str(sp4).upper() if str(sp4) != '#' else None
        self.per4 = int(per4) if str(per4) != '#' else None
        self.sp5 = str(sp5).upper() if str(sp5) != '#' else None
        self.per5 = int(per5) if str(per5) != '#' else None
        self.sp6 = str(sp6).upper() if str(sp6) != '#' else None
        self.per6 = int(per6) if str(per6) != '#' else None
        self.survey_dt = survey_dt
        self.logger = logger

        self.sic_replacement = r'\\bctsdata.bcgov\data\toc_root\Genus_Reporting\GIS_spatial\SIC_Replacement' \
                               r'\SIC_Replacement.gdb\Replacement_Areas'

        self.fld_bec_zone = 'BEC_ZONE_CODE'
        self.fld_bec_subzone = 'BEC_SUBZONE'
        self.fld_bec_var = 'BEC_VARIANT'
        self.fld_age = 'AGE'
        self.fld_dbh = 'DBH'
        self.fld_height = 'HEIGHT'
        self.fld_crown = 'CROWN_CLOSURE'
        self.fld_slope = 'SLOPE'
        self.fld_spec1 = 'SPECIES_1'
        self.fld_perc1 = 'SPECIES_PCT_1'
        self.fld_spec2 = 'SPECIES_2'
        self.fld_perc2 = 'SPECIES_PCT_2'
        self.fld_spec3 = 'SPECIES_3'
        self.fld_perc3 = 'SPECIES_PCT_3'
        self.fld_spec4 = 'SPECIES_4'
        self.fld_perc4 = 'SPECIES_PCT_4'
        self.fld_spec5 = 'SPECIES_5'
        self.fld_perc5 = 'SPECIES_PCT_5'
        self.fld_spec6 = 'SPECIES_6'
        self.fld_perc6 = 'SPECIES_PCT_6'
        self.fld_survey_date = 'SURVEY_DATE'

    def __del__(self):
        arcpy.Delete_management('in_memory')

    def replace_sic(self):
        lst_fields = [self.fld_bec_zone, self.fld_bec_subzone, self.fld_bec_var, self.fld_age, self.fld_dbh,
                      self.fld_height, self.fld_crown, self.fld_slope, self.fld_spec1, self.fld_perc1, self.fld_spec2,
                      self.fld_perc2, self.fld_spec3, self.fld_perc3, self.fld_spec4, self.fld_perc4, self.fld_spec5,
                      self.fld_perc5, self.fld_spec6, self.fld_perc6, self.fld_survey_date, 'SHAPE@']

        with arcpy.da.Editor(os.path.dirname(self.sic_replacement)) as edit:
            with arcpy.da.SearchCursor(self.in_poly, 'SHAPE@') as s_cursor:
                for row in s_cursor:
                    new_shp = row[0].projectAs(arcpy.SpatialReference(3005))
                    with arcpy.da.UpdateCursor(self.sic_replacement, 'SHAPE@') as u_cursor:
                        for u_row in u_cursor:
                            old_shp = u_row[0]
                            if new_shp == old_shp:
                                self.logger.info('New shape is the same as an existing shape, removing old shape')
                                u_cursor.deleteRow()
                                continue
                            elif not new_shp.disjoint(old_shp):
                                self.logger.info('New shape overlaps an existing shape, removing overlap')
                                update_shp = old_shp.difference(new_shp)
                                u_row[0] = update_shp
                                u_cursor.updateRow(u_row)
                                continue

            self.logger.info('Inserting new shape')
            with arcpy.da.InsertCursor(self.sic_replacement, lst_fields) as i_cursor:
                with arcpy.da.SearchCursor(self.in_poly, 'SHAPE@') as s_cursor:
                    for row in s_cursor:
                        new_shp = row[0]
                        new_row = (self.zone, self.sub, self.var, self.age, self.dbh, self.hgt, self.cc, self.slp,
                                   self.sp1, self.per1, self.sp2, self.per2, self.sp3, self.per3, self.sp4, self.per4,
                                   self.sp5, self.per5, self.sp6, self.per6, self.survey_dt, new_shp)
                        i_cursor.insertRow(new_row)


if __name__ == '__main__':
    run_app()
