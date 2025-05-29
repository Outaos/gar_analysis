"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_8006.py

    Author:       BCTS TOC - Graydon Shevchenko
    Purpose:      Class file containing methods specific to GAR 8-006 - Moose
    Date Created: January 10, 2022
----------------------------------------------------------------------------------------------------------------
"""
import xlsxwriter
from datetime import datetime as dt
from util.gar_classes import GARExcel, TotalArea, CellArea
from collections import defaultdict


class Gar8006:
    """
    Class:
        class object for GAR 8-006
    """
    def __init__(self, gar, output_xls, logger, gar_config):
        """
        Function:
            Initializes the Gar8006 object and its attributes
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
        self.str_imm = 'Immature'
        self.str_mature = 'Mature Cover'

        self.str_imm = 'Immature'
        self.lst_level = ['Mature Cover', 'Recruit 1', 'Recruit 2', 'Recruit 3', 'Recruit 4', 'Recruit 5', 'Recruit 6',
                          'Recruit 7', 'Recruit 8', 'Recruit 9', 'Recruit 10', 'Recruit 11']
        self.lst_headers = ['Planning Cell', 'Target (ha) *', 'Analysis Area', 'Total Mature (ha)',
                            'Recruitment - No Harvest (ha)', 'Recruitment - Conditional Harvest (ha)',
                            '+/- (ha) *', 'Total Area (ha) **', 'Mature Cover Over 20 Hectares (Min 50%) ***',
                            'Mature Cover Over 20 Hectares Area (ha) ***',
                            'Mature Cover + Recruitment Over 20 Hectares (Min 50%) ****',
                            'Mature Cover + Recruitment Over 20 Hectares Area (ha) ****',
                            'Immature (Min 15%) *****', 'Immature Area(ha) *****'] + \
                           ['{0} (ha)'.format(i) for i in self.lst_level]

        self.lst_footers = [r'* Positive values show the amount of surplus area within the conditional harvest zones, '
                            r'negative value show a deficit in the Planning Cell.',
                            r'** Total Area is the Gross Forested Land Base not the total area of the cell.',
                            r'*** GAR Schedule 1 - #2: Retain at least 50% of the mature cover requirements in '
                            r'patches of 20 hectares or greater, whenever practical.',
                            r'**** This value is a calculation of any recruitment or mature cover value with '
                            r'harvesting constraints.',
                            r'***** GAR Schedule 1 - #4: Minimum of 15% of the net forested land base of each winter '
                            r'range is to be less that 25 years for ICH and PDF and less then 35 years for MS and '
                            r'ESSF Units. '
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

        if gfa == 'Y' and height and cc:
            if height >= 16:
                if cc >= 56:
                    level = 'Mature Cover'
                elif 46 <= cc < 56:
                    level = 'Recruit 3'
                elif 36 <= cc < 46:
                    level = 'Recruit 6'
                elif 26 <= cc < 36:
                    level = 'Recruit 9'
            elif 14 <= height < 16:
                if cc >= 56:
                    level = 'Recruit 1'
                elif 46 <= cc < 56:
                    level = 'Recruit 4'
                elif 36 <= cc < 46:
                    level = 'Recruit 7'
                elif 26 <= cc < 36:
                    level = 'Recruit 10'
            elif 12 <= height < 14:
                if cc >= 56:
                    level = 'Recruit 2'
                elif 46 <= cc < 56:
                    level = 'Recruit 5'
                elif 36 <= cc < 46:
                    level = 'Recruit 8'
                elif 26 <= cc < 36:
                    level = 'Recruit 11'

            #trying alternative query here as age 0 weren't being classified as immature - daniel 2025-05-13
            if age is not None and ((bec.startswith(('IDF', 'ICH')) and age < 25) or (bec.startswith(('MS', 'ESSF')) and age < 35)):
                level = self.str_imm            

            # if age and ((bec.startswith(('IDF', 'ICH')) and age < 25) or (bec.startswith(('MS', 'ESSF')) and age < 35)):
            #     level = self.str_imm

        if level:
            self.dict_total_area[op_area].pcell[pcell].level[level].hectares += shp_area
            self.dict_cell_area[pcell].level[level].hectares += shp_area

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
                self.dict_total_area[op_area].pcell[pcell].target = cell_area * 0.33

        for pcell in self.dict_cell_area:
            cell_area = self.dict_cell_area[pcell].hectares
            self.dict_cell_area[pcell].target = cell_area * 0.33
        self.calculate_ranks()

    def calculate_ranks(self):
        """
        Function:
            Calls the calculate cell ranks function for both the operating areas and planning cells
        Returns:
            None
        """
        self.logger.info('Calculating ranks')
        self.gar_config.ranks = True
        for op_area in self.dict_total_area:
            self.dict_total_area[op_area].pcell = \
                self.calculate_cell_ranks(dict_cell_area=self.dict_total_area[op_area].pcell)

        self.dict_cell_area = self.calculate_cell_ranks(dict_cell_area=self.dict_cell_area)

    def calculate_cell_ranks(self, dict_cell_area):
        for pcell in dict_cell_area:
            target = dict_cell_area[pcell].target
            running_total = 0
            rank = ''

            for level in self.lst_level:
                level_area = dict_cell_area[pcell].level[level].hectares
                if level_area == 0:
                    continue

                if rank != self.str_ch:
                    running_total += level_area
                    if running_total <= target:
                        rank = self.str_nh
                    else:
                        rank = self.str_ch
                    dict_cell_area[pcell].level[level].rank = rank

                    if level not in [self.str_mature]:
                        if rank == self.str_ch:
                            dict_cell_area[pcell].ch_hectares += level_area
                        elif rank == self.str_nh:
                            dict_cell_area[pcell].nh_hectares += level_area

        return dict_cell_area

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

        for op_area in [o for o in sorted(self.dict_total_area) if o != '']:
            ws = wb.add_worksheet(name=op_area)

            date_now = dt.today().strftime("%B, %Y")
            datestring = 'Created: {}. GAR ORDER: {}'.format(date_now, self.gar)
            ws.write(0, 0, datestring)

            i_row = 1
            i_col = 0
            for c in self.lst_headers:
                style = gar_excel.black_style_bottom_border
                if c in ['Recruitment - No Harvest (ha)']:
                    style = gar_excel.red_style_bottom_border
                elif c in ['Recruitment - Conditional Harvest (ha)']:
                    style = gar_excel.orange_style_bottom_border
                elif c in ['Total Mature (ha)', 'Immature (Min 15%)*****']:
                    style = gar_excel.black_style_bl_border
                elif c in ['Total Area (ha) **', 'Mature Cover Over 20 Hectares Area (ha)***',
                           'Immature Area(ha)*****']:
                    style = gar_excel.black_style_br_border
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
        ch_total = gar_excel.round_value(value=dict_cell_area.ch_hectares)
        nh_total = gar_excel.round_value(value=dict_cell_area.nh_hectares)
        mat_total = gar_excel.round_value(value=dict_cell_area.level[self.str_mature].hectares)
        all_total = mat_total + nh_total + ch_total
        plus_minus = gar_excel.round_value(value=all_total - target)
        total_area = gar_excel.round_value(value=dict_cell_area.hectares)
        mat_pat = gar_excel.round_value(value=dict_cell_area.level[self.str_mature].stand_hectares)
        mat_pat_percent = mat_pat / mat_total if mat_total > 0 else 0
        mat_all = gar_excel.round_value(value=dict_cell_area.stand_hectares)
        mat_all_percent = mat_all / all_total if all_total > 0 else 0
        imm_total = gar_excel.round_value(value=dict_cell_area.level[self.str_imm].hectares)
        imm_percent = imm_total / total_area if imm_total > 0 else 0

        main_style = gar_excel.black_style if analysis == 'Cell' else gar_excel.grey_style
        main_percent_style = gar_excel.black_percent_style if analysis == 'Cell' else gar_excel.grey_percent_style
        main_style_left_border = gar_excel.black_style_left_border if analysis == 'Cell' \
            else gar_excel.grey_style_left_border
        main_style_right_border = gar_excel.black_style_right_border if analysis == 'Cell' \
            else gar_excel.grey_style_right_border
        main_red_style = gar_excel.red_style if analysis == 'Cell' else gar_excel.lite_red_style
        main_orange_style = gar_excel.orange_style if analysis == 'Cell' else gar_excel.lite_orange_style
        main_red_letters = gar_excel.red_letters if analysis == 'Cell' else gar_excel.lite_red_letters
        main_red_letters_percent = gar_excel.red_letters_percent if analysis == 'Cell' \
            else gar_excel.lite_red_letters_percent

        ws.write(i_row, i_col, target, main_style)
        i_col += 1
        ws.write(i_row, i_col, analysis, main_style)
        i_col += 1
        ws.write(i_row, i_col, mat_total, main_style_left_border)
        i_col += 1
        ws.write(i_row, i_col, nh_total, main_red_style)
        i_col += 1
        ws.write(i_row, i_col, ch_total, main_orange_style)
        i_col += 1

        style = main_red_letters if plus_minus < 0 else main_style
        ws.write(i_row, i_col, plus_minus, style)
        i_col += 1
        ws.write(i_row, i_col, total_area, main_style_right_border)
        i_col += 1
        percent_style = main_red_letters_percent if mat_pat_percent < 0.5 else main_percent_style
        ws.write(i_row, i_col, mat_pat_percent, percent_style)
        i_col += 1
        ws.write(i_row, i_col, mat_pat, main_style_right_border)
        i_col += 1
        percent_style = main_red_letters_percent if mat_all_percent < 0.5 else main_percent_style
        ws.write(i_row, i_col, mat_all_percent, percent_style)
        i_col += 1
        ws.write(i_row, i_col, mat_all, main_style_right_border)
        i_col += 1

        percent_style = main_red_letters_percent if imm_percent < 0.15 else main_percent_style
        ws.write(i_row, i_col, imm_percent, percent_style)
        i_col += 1
        ws.write(i_row, i_col, imm_total, main_style_right_border)
        i_col += 1

        for level in level_list:
            rank = dict_cell_area.level[level].rank
            style = main_red_style if rank == 'NH' else main_orange_style if rank == 'CH' else main_style
            if dict_cell_area.level[level].hectares > 0:
                ws.write(i_row, i_col, gar_excel.round_value(value=dict_cell_area.level[level].hectares), style)
            else:
                ws.write(i_row, i_col, '-', style)
            i_col += 1

        return i_col - 1
