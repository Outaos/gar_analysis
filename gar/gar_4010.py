"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_4010.py

    Author:       BCTS TOC - Graydon Shevchenko
    Purpose:      Class file containing methods specific to GAR 4-010 Caribou (Southern Mountain Population)
    Date Created: January 10, 2022
----------------------------------------------------------------------------------------------------------------
"""
import xlsxwriter
from datetime import datetime as dt
from util.gar_classes import GARExcel, TotalArea, CellArea
from collections import defaultdict


class Gar4010:
    """
    Class:
        class object for GAR 4-010
    """
    def __init__(self, gar, output_xls, logger, gar_config):
        """
        Function:
            Initializes the Gar4010 object and its attributes
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

        self.str_no_harv = 'No Harvest Zone'
        self.str_caribou = 'Caribou Management Zone'
        self.str_partial = 'Habitat: Partial Cut'
        self.lst_level = ['Habitat', 'Habitat: Older']
        self.lst_headers = ['Caribou Management Zone', 'Analysis Area']
        str_star = '*'
        for level in ['Habitat', 'Age Class 9']:
            self.lst_headers += ['{0} Target (ha) {1}'.format(level, str_star),
                                 '{0} (ha)'.format(level), '+/- (ha)']
            str_star += '*'
        self.lst_headers += ['Age Class 9 Total Area (%)', 'Total Area (ha) ***']
        self.lst_footers = [r'* Habitat Target is 70% of the total area in cell 5 and 6, 40% of the total '
                            r'area in cell 7 and 30% of the total area in cell 8',
                            r'** Age Class 9 Target is 10% of the total area in cell 7 and 8',
                            r'*** Total Area is the area of the cell where Private Land, Federal Land and Parks have '
                            r'been removed. '
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
        level = None

        if notes == '{0} 1'.format(self.str_caribou):
            level = self.str_no_harv
        else:
            if gfa == 'Y' and slp != '80+':
                if notes in ['{0} 5'.format(self.str_caribou), '{0} 6'.format(self.str_caribou)]:
                    if age and age >= 141:
                        level = self.lst_level[0]
                elif notes == '{0} 7'.format(self.str_caribou):
                    if age and 141 <= age < 251:
                        level = self.lst_level[0]
                    elif age and age >= 251:
                        level = self.lst_level[1]
                elif notes == '{0} 8'.format(self.str_caribou):
                    if age and 141 <= age < 251:
                        level = self.lst_level[0]
                    elif age and age >= 251:
                        level = self.lst_level[1]
                    elif age and 121 <= age < 141:
                        level = self.str_partial

        if level:
            self.dict_total_area[op_area].pcell[notes].level[level].hectares += shp_area
            self.dict_cell_area[notes].level[level].hectares += shp_area

            if level in self.lst_level:
                self.dict_total_area[op_area].pcell[notes].mat_hectares += shp_area
                self.dict_cell_area[notes].mat_hectares += shp_area

        self.dict_total_area[op_area].pcell[notes].hectares += shp_area
        self.dict_cell_area[notes].hectares += shp_area
        if notes not in self.lst_cells:
            self.lst_cells.append(notes)

        if op_area and target == 0:
            self.dict_zero_target[op_area].add(notes)

        return level

    def calculate_targets(self):
        """
        Function:
            Calculates the targets within the operating area an planning cell dictionaries
        Returns:
            None
        """
        for op_area in self.dict_total_area:
            for pcell in self.dict_total_area[op_area].pcell:
                cell_area = self.dict_total_area[op_area].pcell[pcell].hectares
                if pcell in ['{0} 5'.format(self.str_caribou), '{0} 6'.format(self.str_caribou)]:
                    self.dict_total_area[op_area].pcell[pcell].target = cell_area * 0.7
                elif pcell == '{0} 7'.format(self.str_caribou):
                    self.dict_total_area[op_area].pcell[pcell].target = cell_area * 0.4
                elif pcell == '{0} 8'.format(self.str_caribou):
                    self.dict_total_area[op_area].pcell[pcell].target = cell_area * 0.3
                for level in self.dict_total_area[op_area].pcell[pcell].level:
                    if level == self.lst_level[1]:
                        self.dict_total_area[op_area].pcell[pcell].level[level].target = cell_area * 0.1

        for pcell in self.dict_cell_area:
            cell_area = self.dict_cell_area[pcell].hectares
            if pcell in ['{0} 5'.format(self.str_caribou), '{0} 6'.format(self.str_caribou)]:
                self.dict_cell_area[pcell].target = cell_area * 0.7
            elif pcell == '{0} 7'.format(self.str_caribou):
                self.dict_cell_area[pcell].target = cell_area * 0.4
            elif pcell == '{0} 8'.format(self.str_caribou):
                self.dict_cell_area[pcell].target = cell_area * 0.3
            for level in self.dict_cell_area[pcell].level:
                if level == self.lst_level[1]:
                    self.dict_cell_area[pcell].level[level].target = cell_area * 0.1

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
                if c.startswith('Total') or 'Target' in c:
                    style = gar_excel.black_style_bl_border
                ws.write(i_row, i_col, c, style)
                i_col += 1

            end_col = 0

            for pcell in sorted([p for p in self.dict_total_area[op_area].pcell if p != 'Caribou Management Zone 1']):
                i_col = 0
                i_row += 1
                if op_area != '':
                    ws.merge_range(i_row, i_col, i_row + 1, i_col, pcell[-1], gar_excel.black_style_bottom_light_border)
                else:
                    ws.write(i_row, i_col, pcell[-1], gar_excel.black_style)
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
        total_area = gar_excel.round_value(dict_cell_area.hectares)

        main_style = gar_excel.black_style if analysis == 'Cell' else gar_excel.grey_style
        main_percent_style = gar_excel.black_percent_style if analysis == 'Cell' else gar_excel.grey_percent_style
        main_style_left_border = gar_excel.black_style_left_border if analysis == 'Cell' \
            else gar_excel.grey_style_left_border
        main_red_letters = gar_excel.red_letters if analysis == 'Cell' else gar_excel.lite_red_letters
        main_red_letters_percent = gar_excel.red_letters_percent if analysis == 'Cell' \
            else gar_excel.lite_red_letters_percent

        ws.write(i_row, i_col, analysis, main_style)
        i_col += 1
        target = gar_excel.round_value(value=dict_cell_area.target)
        hab = gar_excel.round_value(value=dict_cell_area.mat_hectares)
        plus_minus = hab - target
        ws.write(i_row, i_col, target, main_style_left_border)
        i_col += 1
        ws.write(i_row, i_col, hab, main_style)
        i_col += 1
        if plus_minus < 0:
            style = main_red_letters
        else:
            style = main_style
        ws.write(i_row, i_col, plus_minus, style)
        i_col += 1

        target = gar_excel.round_value(value=dict_cell_area.level[self.lst_level[1]].target)
        hab = gar_excel.round_value(value=dict_cell_area.level[self.lst_level[1]].hectares)
        plus_minus = hab - target
        pct = hab/dict_cell_area.mat_hectares if dict_cell_area.mat_hectares != 0 else 0
        ws.write(i_row, i_col, target, main_style_left_border)
        i_col += 1
        ws.write(i_row, i_col, hab, main_style)
        i_col += 1
        if plus_minus < 0:
            style = main_red_letters
        else:
            style = main_style
        ws.write(i_row, i_col, plus_minus, style)
        i_col += 1
        ws.write(i_row, i_col, pct, main_percent_style)
        i_col += 1
        ws.write(i_row, i_col, total_area, main_style_left_border)

        return i_col
