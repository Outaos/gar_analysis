"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_4001.py

    Author:       BCTS TOC - Graydon Shevchenko
    Purpose:      Class file containing methods specific to GAR 4-001 - Elk, Mule Deer, White Tailed Deer and Moose
    Date Created: January 10, 2022
----------------------------------------------------------------------------------------------------------------
"""

import xlsxwriter
from datetime import datetime as dt
from util.gar_classes import GARExcel, TotalArea, CellArea
from collections import defaultdict


class Gar4001:
    """
    Class:
        class object for GAR 4-001
    """
    def __init__(self, gar, output_xls, logger, gar_config):
        """
        Function:
            Initializes the Gar4001 object and its attributes
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

        self.str_imm = 'Early Seral'
        self.str_mule_deer = 'Mule Deer; ICHmw'
        self.str_moose = 'Moose; moderate snow'
        self.str_foraging = 'Foraging area'
        self.lst_level = ['Mule Deer SIC', 'Moose SIC', 'Forage Areas']
        self.dict_levels = {
            self.str_mule_deer: self.lst_level[0],
            self.str_moose: self.lst_level[1],
            self.str_foraging: self.lst_level[2]
        }
        self.lst_headers = ['Planning Cell', 'Analysis Area']
        str_star = '*'
        for level in self.lst_level:
            self.lst_headers += ['Total {0} (ha)'.format(level), '{0} Target (ha) {1}'.format(level, str_star),
                                 '{0} (ha)'.format(level), '+/- (ha)']
            str_star += '*'
        self.lst_headers += ['Early Seral (ha)', 'Early Seral (max 40%)', 'Total Area (ha) ****']
        self.lst_footers = [r'* Mule Deer SIC Target is 40% of the total SIC area',
                            r'** Moose SIC Target is 20% of the total SIC area',
                            r'*** Forage Area SIC Target is 10% of the total SIC area',
                            r'**** Total Area is the area of the cell where Private Land, Federal Land, Parks, '
                            r'Crown Grants and Broadleaf Stands have been removed.'
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

        if gfa == 'Y':  # If part of the crown forest land base

            # Calculate Mule Deer SIC
            if notes == self.str_mule_deer:
                if (age and cc) and (age >= 101 and cc >= 40):
                    level = self.lst_level[0]

            # Calculate Moose SIC
            elif notes == self.str_moose:
                if (age and cc) and (age >= 61 and cc >= 40):
                    level = self.lst_level[1]

            # Calculate Forage Areas
            elif notes == self.str_foraging:
                if age and age >= 81:
                    level = self.lst_level[2]

            # Calcualte immature
            if age and age < 21:
                level = self.str_imm

            # Add areas to dictionaries
            if level:
                self.dict_total_area[op_area].pcell[pcell].level[level].hectares += shp_area
                self.dict_cell_area[pcell].level[level].hectares += shp_area

            self.dict_total_area[op_area].pcell[pcell].level[self.dict_levels[notes]].total_hectares += shp_area
            self.dict_cell_area[pcell].level[self.dict_levels[notes]].total_hectares += shp_area
            self.dict_cell_area[pcell].hectares += shp_area
            self.dict_total_area[op_area].pcell[pcell].hectares += shp_area
            if pcell not in self.lst_cells:
                self.lst_cells.append(pcell)

            if op_area and target == 0:
                self.dict_zero_target[op_area].add(pcell)

        return level

    def calculate_targets(self):
        """
        Function:
            Calculates the targets within the operating area an planning cell dictionaries
        """
        for op_area in self.dict_total_area:
            for pcell in self.dict_total_area[op_area].pcell:
                for level in self.dict_total_area[op_area].pcell[pcell].level:
                    if level == self.lst_level[0]:  # Mule Deer: SIC
                        self.dict_total_area[op_area].pcell[pcell].level[level].target = \
                            self.dict_total_area[op_area].pcell[pcell].level[level].total_hectares * 0.4
                    elif level == self.lst_level[1]:  # Moose: SIC
                        self.dict_total_area[op_area].pcell[pcell].level[level].target = \
                            self.dict_total_area[op_area].pcell[pcell].level[level].total_hectares * 0.2
                    elif level == self.lst_level[2]:  # Forage Areas
                        self.dict_total_area[op_area].pcell[pcell].level[level].target = \
                            self.dict_total_area[op_area].pcell[pcell].level[level].total_hectares * 0.1

        for pcell in self.dict_cell_area:
            for level in self.dict_cell_area[pcell].level:
                if level == self.lst_level[0]:  # Mule Deer: SIC
                    self.dict_cell_area[pcell].level[level].target = \
                        self.dict_cell_area[pcell].level[level].total_hectares * 0.4
                elif level == self.lst_level[1]:  # Moose: SIC
                    self.dict_cell_area[pcell].level[level].target = \
                        self.dict_cell_area[pcell].level[level].total_hectares * 0.2
                elif level == self.lst_level[2]:  # Forage Areas
                    self.dict_cell_area[pcell].level[level].target = \
                        self.dict_cell_area[pcell].level[level].total_hectares * 0.1

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
                if c.startswith(('Total', 'Early Seral (ha)')):
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
                                           i_row=i_row, i_col=i_col, analysis='Cell', gar_excel=gar_excel)
                if op_area != '':
                    i_row += 1
                    end_col = self.write_cells(dict_cell_area=self.dict_total_area[op_area].pcell[pcell], ws=ws,
                                               i_row=i_row, i_col=i_col, analysis='Op Area', gar_excel=gar_excel)

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

    def write_cells(self, dict_cell_area, ws, i_row, i_col, analysis, gar_excel):
        """
        Function:
            Loops through the dictionaries and writes the analysis values to the cells in the Excel document
        Args:
            dict_cell_area (CellArea):
            ws:
            i_row:
            i_col:
            analysis:
            gar_excel:

        Returns:

        """
        total_area = gar_excel.round_value(dict_cell_area.hectares)
        early_seral = gar_excel.round_value(value=dict_cell_area.level[self.str_imm].hectares)
        early_seral_percent = early_seral / total_area if total_area != 0 else 0

        main_style = gar_excel.black_style if analysis == 'Cell' else gar_excel.grey_style
        main_percent_style = gar_excel.black_percent_style if analysis == 'Cell' else gar_excel.grey_percent_style
        main_style_left_border = gar_excel.black_style_left_border if analysis == 'Cell' \
            else gar_excel.grey_style_left_border
        main_red_letters = gar_excel.red_letters if analysis == 'Cell' else gar_excel.lite_red_letters
        main_red_letters_percent = gar_excel.red_letters_percent if analysis == 'Cell' \
            else gar_excel.lite_red_letters_percent

        ws.write(i_row, i_col, analysis, main_style)
        i_col += 1
        for level in self.lst_level:
            total_sic = gar_excel.round_value(value=dict_cell_area.level[level].total_hectares)
            target = gar_excel.round_value(value=dict_cell_area.level[level].target)
            sic = gar_excel.round_value(value=dict_cell_area.level[level].hectares)
            plus_minus = sic - target
            ws.write(i_row, i_col, total_sic, main_style_left_border)
            i_col += 1
            ws.write(i_row, i_col, target, main_style)
            i_col += 1
            ws.write(i_row, i_col, sic, main_style)
            i_col += 1
            if plus_minus < 0:
                style = main_red_letters
            else:
                style = main_style
            ws.write(i_row, i_col, plus_minus, style)
            i_col += 1

        ws.write(i_row, i_col, early_seral, main_style_left_border)
        i_col += 1
        percent_style = main_red_letters_percent if early_seral_percent > 0.4 else main_percent_style
        ws.write(i_row, i_col, early_seral_percent, percent_style)
        i_col += 1
        ws.write(i_row, i_col, total_area, main_style_left_border)

        return i_col
