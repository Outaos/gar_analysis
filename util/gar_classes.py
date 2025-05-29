"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_classes.py

    Author:       BCTS TOC - Graydon Shevchenko
    Purpose:      Contains classes used by the GAR Analysis
    Date Created: January 10, 2022
----------------------------------------------------------------------------------------------------------------
"""
from collections import defaultdict


class TotalArea:
    """
    Class:
        Object used to track planning cells within each operating area
    """
    def __init__(self):
        self.pcell = defaultdict(CellArea)


class CellArea:
    """
    Class:
        Object used to track levels within each planning cell
    """
    def __init__(self):
        self.level = defaultdict(self.Level)
        self.hectares = 0
        self.target = 0
        self.nthlb_hectares = 0
        self.sic_hectares = 0
        self.ch_hectares = 0
        self.nh_hectares = 0
        self.imm_hectares = 0
        self.mat_hectares = 0
        self.stand_hectares = 0

    class Level:
        """
        Class:
            Object used to track bec within each level
        """
        def __init__(self):
            self.bec = defaultdict(self.BEC)
            self.hectares = 0
            self.total_hectares = 0
            self.rank = None
            self.target = 0
            self.stand_hectares = 0

        class BEC:
            """
            Class:
                Object used to hold bec area and rank
            """
            def __init__(self):
                self.hectares = 0
                self.rank = None


class GARConfig:
    """
    Class:
        Configuration class object
    """
    def __init__(self, sql=None, cells=None, cell_field=None, aoi=None, private_land=None, erase_fcs=None,
                 identity_fcs=None, ranks=False):
        self.sql = sql
        self.cells = cells
        self.cell_field = cell_field
        self.aoi = aoi
        self.private_land = private_land
        self.erase_fcs = erase_fcs if erase_fcs else []
        self.identity_fcs = identity_fcs if identity_fcs else []
        self.ranks = ranks


class GARInput:
    """
    Class:
        Input class object
    """
    def __init__(self, path=None, sql=None, output=None, mandatory=False):
        self.path = path
        self.sql = sql
        self.output = output
        self.mandatory = mandatory

class SICReplacement:
    """
    Class:
        SIC replacement class object
    """

    def __init__(self, zone=None, sub=None, var=None, age=None, dbh=None, hgt=None, cc=None, slp=None,
                 sp1=None, per1=None, sp2=None, per2=None, sp3=None, per3=None, sp4=None, per4=None, sp5=None,
                 per5=None, sp6=None, per6=None, survey_dt=None):
        self.zone = zone
        self.sub = sub
        self.var = var
        self.age = age
        self.dbh = dbh
        self.hgt = hgt
        self.cc = cc
        self.slp = slp
        self.sp1 = sp1
        self.per1 = per1
        self.sp2 = sp2
        self.per2 = per2
        self.sp3 = sp3
        self.per3 = per3
        self.sp4 = sp4
        self.per4 = per4
        self.sp5 = sp5
        self.per5 = per5
        self.sp6 = sp6
        self.per6 = per6
        self.survey_dt = survey_dt


class GARExcel:
    """
    Class:
        Excel class object containing formatting used in xlsxwriter
    """
    def __init__(self, wb):
        self.wb = wb
        self.orange_style = wb.add_format({'font_color': '#9c6500', 'bg_color': '#ffeb9c', 'text_wrap': True})
        self.lite_orange_style = \
            wb.add_format({'font_color': '#b87b00', 'bg_color': '#fff3c1', 'text_wrap': True, 'bottom': 1})
        self.red_style = wb.add_format({'font_color': '#9c0006', 'bg_color': '#ffc7ce', 'text_wrap': True})
        self.lite_red_style = \
            wb.add_format({'font_color': '#ff9398', 'bg_color': '#ffe5e8', 'text_wrap': True, 'bottom': 1})
        self.red_letters = wb.add_format({'font_color': '#ff0000', 'text_wrap': True, 'bold': True})
        self.lite_red_letters = wb.add_format({'font_color': '#ff9398', 'text_wrap': True, 'bold': True, 'bottom': 1})
        self.red_letters_percent = \
            wb.add_format({'num_format': '0%', 'font_color': '#ff0000', 'text_wrap': True, 'bold': True})
        self.lite_red_letters_percent = \
            wb.add_format({'num_format': '0%', 'font_color': '#ff9398', 'text_wrap': True, 'bold': True, 'bottom': 1})
        self.black_style = wb.add_format({'font_color': '#000000', 'text_wrap': True})
        self.grey_style = wb.add_format({'font_color': '#666666', 'text_wrap': True, 'bottom': 1})
        self.black_percent_style = wb.add_format({'num_format': '0%', 'font_color': '#000000', 'text_wrap': True})
        self.grey_percent_style = \
            wb.add_format({'num_format': '0%', 'font_color': '#666666', 'text_wrap': True, 'bottom': 1})
        self.red_style_bottom_border = \
            wb.add_format({'font_color': '#9c0006', 'bg_color': '#ffc7ce', 'text_wrap': True, 'bottom': 2})
        self.black_style_bottom_border = wb.add_format({'font_color': '#000000', 'text_wrap': True, 'bottom': 2})
        self.black_style_bottom_light_border = wb.add_format({'font_color': '#000000', 'text_wrap': True, 'bottom': 1})
        self.black_style_top_light_border = wb.add_format({'font_color': '#000000', 'text_wrap': True, 'top': 1})
        self.orange_style_bottom_border = \
            wb.add_format({'font_color': '#9c6500', 'bg_color': '#ffeb9c', 'text_wrap': True, 'bottom': 2})
        self.black_style_left_border = wb.add_format({'font_color': '#000000', 'text_wrap': True, 'left': 6})
        self.grey_style_left_border = \
            wb.add_format({'font_color': '#666666', 'text_wrap': True, 'left': 6, 'bottom': 1})
        self.black_style_right_border = wb.add_format({'font_color': '#000000', 'text_wrap': True, 'right': 6})
        self.grey_style_right_border = \
            wb.add_format({'font_color': '#666666', 'text_wrap': True, 'right': 6, 'bottom': 1})
        self.black_style_bl_border = wb.add_format({'font_color': '#000000', 'text_wrap': True, 'bottom': 2, 'left': 6})
        self.black_style_br_border = \
            wb.add_format({'font_color': '#000000', 'text_wrap': True, 'bottom': 2, 'right': 6})

    def round_value(self, value):
        """
        Function:
            Rounds an input value to 2 decimal places; values with more than 1 zero preceding a number will be
            rounded to the appropriate number of digits to show a non zero number
        Args:
            value (float): input value to be rounded

        Returns:
            float: the resultant rounded value

        """
        if isinstance(value, float):
            value_decimal = str(value).split('.')[1]
            i = 1
            for j in range(0, len(value_decimal)):
                if value_decimal[j] == '0':
                    i += 1
                    continue
                else:
                    break
            if i > 2:
                return round(value, i)
            else:
                return round(value, 2)
        return value
