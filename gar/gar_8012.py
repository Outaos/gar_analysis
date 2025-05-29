"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_8012.py

    Author:       BCTS TOC - Graydon Shevchenko
    Purpose:      Class file containing methods specific to GAR 8-012 - Caribou (Southern Mountain Population)
    Date Created: January 10, 2022
----------------------------------------------------------------------------------------------------------------
"""
import xlsxwriter
from datetime import datetime as dt
from util.gar_classes import GARExcel, TotalArea, CellArea
from collections import defaultdict


class Gar8012:
    """
    Class:
        class object for GAR 8-012
    """
    def __init__(self, gar, output_xls, logger, gar_config):
        """
        Function:
            Initializes the Gar8012 object and its attributes
        Args:
            gar (str): input gar string
            output_xls (str): output excel file path
            logger (logger): logger object
            gar_config (GARConfig): configuration class object applicable to this gar
        Returns:
            None
        """
        self.gar = gar
        self.output_xls = output_xls
        self.logger = logger
        self.gar_config = gar_config
        self.dict_total_area = defaultdict(TotalArea)
        self.dict_cell_area = defaultdict(CellArea)
        self.dict_target = defaultdict(float)
        self.dict_zero_target = defaultdict(set)
        self.lst_cells = []

        self.str_ch = 'CH'
        self.str_nh = 'NH'
        self.str_suitable = 'Suitable Cover'
        self.str_no_harv = 'No Harvest'

        self.str_imm = 'Immature'
        self.lst_level = [self.str_suitable]
        self.lst_headers = ['Planning Cell', 'Suitable Cover Target (ha) *', 'Analysis Area', 'Suitable Cover (ha)',
                            '+/- (ha) *', 'Total Area (ha) **']

        self.lst_footers = [r'* Suitable Cover Target is 70% of the total cell area',
                            r'** Total Area is the area of the cell where Private Land has been removed.'
                            ]

    def calculate_level(self, bec, age, spp, cc, slp, thlb, diam, pct, gfa,
                        notes, op_area, pcell, shp_area, target, height):
        """
        Function:
            Calculates the levels applicable to this gar
        Args:
            bec (str): bec
            age (int): age
            spp (str): species
            cc (int): crown closure
            slp (str): slope - None or 80+
            thlb (float): timber harvest landbase
            diam (float): diameter
            pct (int): species percent
            gfa (str): gross forested area - None or Y
            notes (str): feature notes
            op_area (str): operating area name
            pcell (str): planning cell
            shp_area (float): feature area
            target (float): target
            height (float): tree height

        Returns:
            str: level assigned to record
        """
        bec = bec.replace(' ', '')
        level = None
        if age and bec.startswith(('ESSF', 'ICH')) and age >= 141:
            level = 'Suitable Cover'
        if bec.startswith('ESSFwcp'):
            level = self.str_no_harv

        if level and level != self.str_no_harv:
            self.dict_total_area[op_area].pcell[pcell].level[level].hectares += shp_area
            self.dict_cell_area[pcell].level[level].hectares += shp_area

        if level != self.str_no_harv and bec.startswith(('ESSF', 'ICH')):
            self.dict_total_area[op_area].pcell[pcell].hectares += shp_area
            self.dict_cell_area[pcell].hectares += shp_area

        if pcell not in self.lst_cells:
            self.lst_cells.append(pcell)

        if op_area and target == 0:
            self.dict_zero_target[op_area].add(pcell)

        return level

    def calculate_targets(self):
        """
        Function:
            Calculates the targets within the operating area an planning cell dictionaries
        Returns:
            None
        """
        self.logger.info('Calculating targets')
        for op_area in self.dict_total_area:
            for pcell in self.dict_total_area[op_area].pcell:
                cell_area = self.dict_total_area[op_area].pcell[pcell].hectares
                self.dict_total_area[op_area].pcell[pcell].target = cell_area * 0.7

        for pcell in self.dict_cell_area:
            cell_area = self.dict_cell_area[pcell].hectares
            self.dict_cell_area[pcell].target = cell_area * 0.7

    def write_excel(self):
        """
        Function:
            Writes the analysis results to an Excel file using Xlsxwriter
        Returns:
            None
        """
        self.logger.info('Writing to excel')
        wb = xlsxwriter.Workbook(filename=self.output_xls)
        gar_excel = GARExcel(wb=wb)

        for op_area in [o for o in sorted(self.dict_total_area)]:
            if op_area == '':
                ws = wb.add_worksheet(name='No Operating Area')
            else:
                ws = wb.add_worksheet(name=op_area)

            date_now = dt.today().strftime("%B, %Y")
            datestring = 'Created: {}. GAR ORDER: {}'.format(date_now, self.gar)
            ws.write(0, 0, datestring)

            i_row = 1
            i_col = 0
            for c in self.lst_headers:
                style = gar_excel.black_style_bottom_border
                if c.startswith(('Total', 'Suitable Cover (ha)')):
                    style = gar_excel.black_style_bl_border
                ws.write(i_row, i_col, c, style)
                i_col += 1

            end_col = 0
            for pcell in sorted([p for p in self.dict_total_area[op_area].pcell]):
                i_col = 0
                i_row += 1
                if op_area != '':
                    ws.merge_range(i_row, i_col, i_row + 1, i_col, pcell, gar_excel.black_style_bottom_light_border)
                else:
                    ws.write(i_row, i_col, pcell, gar_excel.black_style)
                i_col += 1
                end_col = self.write_cells(dict_cell_area=self.dict_cell_area[pcell], ws=ws,
                                           i_row=i_row, i_col=i_col, analysis='Cell', level_list=self.lst_level,
                                           gar_excel=gar_excel)
                if op_area != '':
                    i_row += 1
                    end_col = self.write_cells(dict_cell_area=self.dict_total_area[op_area].pcell[pcell], ws=ws,
                                               i_row=i_row, i_col=i_col, analysis='Op Area', level_list=self.lst_level,
                                               gar_excel=gar_excel)

            end_row = i_row + 1
            for line in self.lst_footers:
                if line == self.lst_footers[0]:
                    style = gar_excel.black_style_top_light_border
                else:
                    style = gar_excel.black_style
                i_row += 1
                line_length = len(line)
                num_lines = 1
                while line_length > 115:
                    line_length -= 115
                    num_lines += 1
                ws.merge_range(i_row, 0, i_row, 11, line, style)
                ws.set_row(i_row, 15 * num_lines)

            i_col = 12
            while i_col <= end_col:
                ws.write(end_row, i_col, None, gar_excel.black_style_top_light_border)
                i_col += 1

        wb.close()

    def write_cells(self, dict_cell_area, ws, i_row, i_col, analysis, level_list, gar_excel):
        target = gar_excel.round_value(value=dict_cell_area.target)
        sc_total = gar_excel.round_value(value=dict_cell_area.level[self.str_suitable].hectares)
        plus_minus = gar_excel.round_value(value=sc_total - target)
        total_area = gar_excel.round_value(value=dict_cell_area.hectares)

        main_style = gar_excel.black_style if analysis == 'Cell' else gar_excel.grey_style
        main_style_left_border = gar_excel.black_style_left_border if analysis == 'Cell' \
            else gar_excel.grey_style_left_border

        main_red_letters = gar_excel.red_letters if analysis == 'Cell' else gar_excel.lite_red_letters

        ws.write(i_row, i_col, target, main_style)
        i_col += 1
        ws.write(i_row, i_col, analysis, main_style)
        i_col += 1
        ws.write(i_row, i_col, sc_total, main_style_left_border)
        i_col += 1
        style = main_red_letters if plus_minus < 0 else main_style
        ws.write(i_row, i_col, plus_minus, style)
        i_col += 1
        ws.write(i_row, i_col, total_area, main_style_left_border)
        i_col += 1

        return i_col - 1
