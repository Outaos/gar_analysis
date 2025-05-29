"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_8001.py

    Author:       BCTS TOC - Graydon Shevchenko
    Purpose:      Class file containing methods specific to GAR 8-001 - Mule Deer
    Date Created: January 10, 2022
----------------------------------------------------------------------------------------------------------------
"""
import xlsxwriter
from datetime import datetime as dt
from util.gar_classes import GARExcel, TotalArea, CellArea
from collections import defaultdict


class Gar8001:
    """
    Class:
        class object for GAR 8-001
    """
    def __init__(self, gar, output_xls, logger, gar_config):
        """
        Function:
            Initializes the Gar8001 object and its attributes
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
        self.str_nthlb_sic = 'NTHLB SIC'
        self.str_sic = 'SIC'

        self.str_imm = 'Immature'
        self.lst_level = ['NTHLB SIC', 'SIC', 'NTHLB Recruit 1', 'Recruit 1', 'NTHLB Recruit 2', 'Recruit 2',
                          'NTHLB Recruit 3', 'Recruit 3', 'NTHLB Recruit 4', 'Recruit 4', 'NTHLB Recruit 5',
                          'Recruit 5', 'NTHLB Recruit 6', 'Recruit 6', 'NTHLB Recruit 7', 'Recruit 7',
                          'NTHLB Recruit 8', 'Recruit 8', 'NTHLB Recruit 9', 'Recruit 9', 'NTHLB Recruit 10',
                          'Recruit 10', 'NTHLB Recruit 11', 'Recruit 11', 'NTHLB Recruit 12', 'Recruit 12']
        self.lst_headers = ['Planning Cell', 'Target (ha) *', 'Analysis Area', 'Total SIC (ha)',
                            'Recruitment - No Harvest (ha)', 'Recruitment - Conditional Harvest (ha)',
                            '+/- (ha) **', 'Total Area (ha) ***', 'Mod Snow Pack Immature (Max 30%) ****',
                            'Mod Snow Pack Immature Area(ha)', 'Mod Snow Pack NTHLB SIC (Max 50%) *****',
                            'Mod Snow Pack NTHLB SIC Area(ha)'] + \
                           ['{0} (ha)'.format(i) for i in self.lst_level]
        self.lst_footers = [r'* The targets of the Operating Area analysis are proportional.',
                            r'** Positive values show a surplus area within the conditional harvest zones, '
                            r'negative value show a deficit in the Planning Cell.',
                            r'*** Total Area is the area of the cell where Private Land and Woodlots have been '
                            r'removed.',
                            r'**** GAR Schedule 1 - #12: In the Moderate Snowpack Zone, no more than 30% of the '
                            r'planning cell is to be in stands of less than 20 years of age.',
                            r'***** GAR Schedule 1 - #6: In the Moderate Snow Pack Zone (except IDFmw up to 50% '
                            r'of the Snow Interception Cover can be met in the NTHLB,provided the stands are '
                            r'at least 50% Douglas-Fir, at least 120 year of age, and have a crown closure of at '
                            r'least 36% '
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
        thlb = float(thlb) if thlb else 0
        level = None

        # ================================================================
        #  SHALLOW Snowpack Zone
        # ================================================================
        if bec.startswith(('BG', 'PP', 'IDFxh')) and (spp.startswith('F')):
            if age:
                if age >= 140:
                    level = 'SIC'
                elif 130 <= age <= 139:
                    level = 'Recruit 1'
                elif 120 <= age <= 129:
                    level = 'Recruit 2'
                elif 110 <= age <= 119:
                    level = 'Recruit 3'
                elif 100 <= age <= 109:
                    level = 'Recruit 4'
                elif 90 <= age <= 99:
                    level = 'Recruit 5'
                elif 80 <= age <= 89:
                    level = 'Recruit 6'
                elif 70 <= age <= 79:
                    level = 'Recruit 7'
                elif 60 <= age <= 69:
                    level = 'Recruit 8'

        # ================================================================
        #  MODERATE Snowpack Zone
        # ================================================================
        # IDFmw is handled different than other BEC's
        elif bec.startswith('IDFmw'):
            # any IDFmw with Fir Leading, Crown Closure greater than 36, not on slope greater than 80%, and a THLB
            # greater than 0.
            if spp.startswith('F') and slp != '80+':
                if thlb > 0:
                    if cc and cc >= 36:
                        # SIC is any stands with age greater than 140 or a diameter greater then 40 and age over 0.
                        # As per dicussions with M. Kyler anywhere diam > 40cm it is automatically added to THLB SIC
                        # and disregards the age of poly.
##                        if (age and age >= 140) or (diam >= 40 and age > 0):
## replaced line above with one below to remove error encountered 2023/04/26 Daniel
                        if (age is not None and age >= 140) or (diam is not None and diam >= 40 and age is not None and age > 0):
                            level = 'SIC'
##                        if diam < 40:
## replaced line above with one below to remove error encountered 2023/04/26 Daniel
                        if diam is not None and diam < 40:
                            if age:
                                if 130 <= age <= 139:
                                    level = 'Recruit 1'
                                elif 120 <= age <= 129:
                                    level = 'Recruit 2'
                                elif 110 <= age <= 119:
                                    level = 'Recruit 3'
                                elif 100 <= age <= 109:
                                    level = 'Recruit 4'
                                elif 90 <= age <= 99:
                                    level = 'Recruit 5'
                                elif 80 <= age <= 89:
                                    level = 'Recruit 6'
                                elif 70 <= age <= 79:
                                    level = 'Recruit 7'
                                elif 60 <= age <= 69:
                                    level = 'Recruit 8'

                # any IDFmw with Fir Leading, Crown Closure greater than 50, not on slope greater than 80%,
                # and a THLB = 0.
                elif thlb == 0:
                    if (cc and age and pct) and cc >= 50 and age >= 120 and pct >= 50:
                        level = 'NTHLB SIC'
                    elif cc and cc >= 36:
                        # Recruit 1 and 2 must have a cc >= 36 and cc <50 to not over write the NTHLB SIC
                        if (36 <= cc < 50) and (130 <= age <= 139):
                            level = 'NTHLB Recruit 1'
                        elif (36 <= cc < 50) and (120 <= age <= 129):
                            level = 'NTHLB Recruit 2'
                        elif 110 <= age <= 119:
                            level = 'NTHLB Recruit 3'
                        elif 100 <= age <= 109:
                            level = 'NTHLB Recruit 4'
                        elif 90 <= age <= 99:
                            level = 'NTHLB Recruit 5'
                        elif 80 <= age <= 89:
                            level = 'NTHLB Recruit 6'
                        elif 70 <= age <= 79:
                            level = 'NTHLB Recruit 7'
                        elif 60 <= age <= 69:
                            level = 'NTHLB Recruit 8'

        elif bec.startswith(('IDFdk', 'IDFdm', 'ICHdw', 'MS')):
            if (spp.startswith('F')) and (cc >= 36) and (slp != '80+'):
                if thlb > 0:
                    # calculate the SIC levels As per dicussions with M. Kyler anywhere diam > 40cm it is
                    # automatically added to THLB SIC and disregards the age of poly.
##                    if (age and age >= 175) or (diam >= 40 and age > 0):
## replaced line above with one below to remove error encountered 2023/04/26 Daniel
                    if (age is not None and age >= 175) or (diam is not None and diam >= 40 and age is not None and age > 0):
                        level = 'SIC'
                    # adding less than 40 diam so recruit values dont overwrite sic values
##                    if diam and diam < 40:
## replaced line above with one below to remove error encountered 2023/04/26 Daniel
                    if diam is not None and diam < 40:
                        if age:
                            if 165 <= age <= 174:
                                level = 'Recruit 1'
                            elif 155 <= age <= 164:
                                level = 'Recruit 2'
                            elif 145 <= age <= 154:
                                level = 'Recruit 3'
                            elif 135 <= age <= 144:
                                level = 'Recruit 4'
                            elif 125 <= age <= 134:
                                level = 'Recruit 5'
                            elif 115 <= age <= 124:
                                level = 'Recruit 6'
                            elif 105 <= age <= 114:
                                level = 'Recruit 7'
                            elif 95 <= age <= 104:
                                level = 'Recruit 8'
                            elif 85 <= age <= 94:
                                level = 'Recruit 9'
                            elif 75 <= age <= 84:
                                level = 'Recruit 10'
                            elif 65 <= age <= 74:
                                level = 'Recruit 11'
                            elif 55 <= age <= 64:
                                level = 'Recruit 12'

                elif thlb == 0:
                    if age:
                        if (age >= 120) and (pct >= 50):
                            level = 'NTHLB SIC'
                        # ***  Recruit 6 / NTHLB Recruit 6 (The first recruitment level for the NTHLB data as >=120yrs is
                        # SIC so can't recruit until this age level; >=115...)  ***

                        elif 115 <= age <= 119:
                            level = 'NTHLB Recruit 6'
                        elif 105 <= age <= 114:
                            level = 'NTHLB Recruit 7'
                        elif 95 <= age <= 104:
                            level = 'NTHLB Recruit 8'
                        elif 85 <= age <= 94:
                            level = 'NTHLB Recruit 9'
                        elif 75 <= age <= 84:
                            level = 'NTHLB Recruit 10'
                        elif 65 <= age <= 74:
                            level = 'NTHLB Recruit 11'
                        elif 55 <= age <= 64:
                            level = 'NTHLB Recruit 12'

        # ================================================================
        #  DEEP Snowpack Zone
        # ================================================================
        elif bec.startswith('ICH') and not bec.startswith('ICHdw'):
            if spp.startswith('F'):
                if cc and cc >= 46:
                    if age and diam:
                        if age >= 100 or (diam >= 40 and age > 0):
                            level = 'SIC'
                        elif (90 <= age <= 99) and (diam < 40):
                            level = 'Recruit 1'
                        elif (80 <= age <= 89) and (diam < 40):
                            level = 'Recruit 2'
                        elif (70 <= age <= 79) and (diam < 40):
                            level = 'Recruit 6'
                        elif (60 <= age <= 79) and (diam < 40):
                            level = 'Recruit 8'
                elif cc and 36 <= cc < 46:
                    if age and diam:
                        if age >= 100 or (diam >= 40 and age > 0):
                            level = 'Recruit 3'
                        elif (90 <= age <= 99) and (diam < 40):
                            level = 'Recruit 4'
                        elif (80 <= age <= 89) and (diam < 40):
                            level = 'Recruit 5'
                        elif (70 <= age <= 79) and (diam < 40):
                            level = 'Recruit 7'
                        elif (60 <= age <= 69) and (diam < 40):
                            level = 'Recruit 9'

        if age and age < 20:
            if bec.startswith(('IDFdk', 'IDFdm', 'ICHdw', 'MS', 'IDFmw')):
                level = self.str_imm

        self.dict_total_area[op_area].pcell[pcell].level[level].bec[bec].hectares += shp_area
        self.dict_cell_area[pcell].level[level].bec[bec].hectares += shp_area
        self.dict_cell_area[pcell].target = target
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
        Returns:
            None
        """
        self.logger.info('Calculating targets')
        for op_area in self.dict_total_area:
            for pcell in self.dict_total_area[op_area].pcell:
                self.dict_total_area[op_area].pcell[pcell].target = \
                    self.dict_cell_area[pcell].target * \
                    (self.dict_total_area[op_area].pcell[pcell].hectares / self.dict_cell_area[pcell].hectares)
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
            target_50 = target * 0.5
            running_total = 0
            rank = ''

            for level in self.lst_level + [self.str_imm]:
                for bec in dict_cell_area[pcell].level[level].bec:
                    bec_area = dict_cell_area[pcell].level[level].bec[bec].hectares
                    if level == self.str_nthlb_sic:
                        if not bec.startswith('IDFmw') and bec_area >= target_50:
                            dict_cell_area[pcell].level[level].bec[bec].hectares = target_50
                            bec_area = dict_cell_area[pcell].level[level].bec[bec].hectares
                        if bec.startswith(('IDFdk', 'IDFdm', 'ICHdw', 'MS')):
                            dict_cell_area[pcell].nthlb_hectares += bec_area
                        dict_cell_area[pcell].sic_hectares += bec_area
                    elif level == self.str_sic:
                        dict_cell_area[pcell].sic_hectares += bec_area
                    elif level == self.str_imm:
                        dict_cell_area[pcell].imm_hectares += bec_area
                        continue

                    if rank != self.str_ch:
                        running_total += bec_area
                        if running_total <= target:
                            rank = self.str_nh
                        else:
                            rank = self.str_ch
                        dict_cell_area[pcell].level[level].bec[bec].rank = rank
                        dict_cell_area[pcell].level[level].rank = rank

                        if level not in [self.str_nthlb_sic, self.str_sic]:
                            if rank == self.str_ch:
                                dict_cell_area[pcell].ch_hectares += bec_area
                            elif rank == self.str_nh:
                                dict_cell_area[pcell].nh_hectares += bec_area
                    dict_cell_area[pcell].level[level].hectares += bec_area

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

        for op_area in [o for o in sorted(self.dict_total_area)]:
            # #Change this part so it runs on all Op Areas and non operaitng areas - Daniel 2025/03/24
            # if self.gar != 'u-8-001-tfl49' and op_area == '':
            #      continue
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
                if c in ['Recruitment - No Harvest (ha)']:
                    style = gar_excel.red_style_bottom_border
                elif c in ['Recruitment - Conditional Harvest (ha)']:
                    style = gar_excel.orange_style_bottom_border
                elif c in ['Total SIC (ha)', 'Mod Snow Pack Immature (Max 30%) ****']:
                    style = gar_excel.black_style_bl_border
                elif c in ['Mod Snow Pack Immature Area(ha)', 'Mod Snow Pack NTHLB SIC Area(ha)']:
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
        sic_total = gar_excel.round_value(value=dict_cell_area.sic_hectares)
        total_area = gar_excel.round_value(value=dict_cell_area.hectares)
        imm_total = gar_excel.round_value(value=dict_cell_area.imm_hectares)
        imm_percent = imm_total / total_area if imm_total > 0 else 0
        plus_minus = gar_excel.round_value(value=(sic_total + nh_total + ch_total) - target)
        nthlb_total = gar_excel.round_value(value=dict_cell_area.nthlb_hectares)
        nthlb_percent = nthlb_total / target if target > 0 else 0
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
        ws.write(i_row, i_col, sic_total, main_style_left_border)
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
        percent_style = main_red_letters_percent if imm_percent > 0.3 else main_percent_style
        ws.write(i_row, i_col, imm_percent, percent_style)
        i_col += 1
        ws.write(i_row, i_col, imm_total, main_style_right_border)
        i_col += 1
        percent_style = main_red_letters_percent if nthlb_percent > 0.5 else main_percent_style
        ws.write(i_row, i_col, nthlb_percent, percent_style)
        i_col += 1
        ws.write(i_row, i_col, nthlb_total, main_style_right_border)
        i_col += 1

        for level in self.lst_level:
            rank = dict_cell_area.level[level].rank
            style = main_red_style if rank == 'NH' else main_orange_style if rank == 'CH' else main_style
            if dict_cell_area.level[level].hectares > 0:
                ws.write(i_row, i_col, gar_excel.round_value(value=dict_cell_area.level[level].hectares), style)
            else:
                ws.write(i_row, i_col, '-', style)
            i_col += 1

        return i_col - 1
