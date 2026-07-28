"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_8008.py

    Purpose:      Spatial screening class for GAR U-8-008 - Mule Deer.
    Source:       https://www.env.gov.bc.ca/wld/documents/uwr/U-8-008_ord.pdf

    Implemented legal measures:
      * cell-specific snow interception cover (SIC) targets from Schedule 1, Table 1;
      * SIC age and canopy thresholds by snowpack zone from Tables 2 and 3;
      * at least 67% of the Moderate Snowpack Zone over 20 years of age; and
      * open-road density not exceeding 3 km/km2 when road lengths are supplied.

    The current generic GAR cursor does not supply Fire Maintained Ecosystem Restoration (FMER), elevation,
    aspect, road-length or proposed-access attributes. Optional arguments and record_open_road_length() allow
    those measures to be populated later. Until then, the Excel report labels affected results as incomplete or
    provisional instead of claiming full compliance.
----------------------------------------------------------------------------------------------------------------
"""

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime as dt
import re

import xlsxwriter

from util.gar_classes import GARExcel, TotalArea, CellArea


@dataclass
class U8008CellStats:
    """Areas and access values accumulated for one U-8-008 planning cell."""

    gross_hectares: float = 0.0
    net_hectares: float = 0.0
    sic_hectares: float = 0.0
    moderate_hectares: float = 0.0
    moderate_age_over_20_hectares: float = 0.0
    fmer_unassessed_hectares: float = 0.0
    unresolved_snowpack_hectares: float = 0.0
    open_road_km: float = 0.0
    road_density_assessed: bool = False
    access_in_sic_conflict: bool = False


class Gar8008:
    """Spatial screening and Excel reporting for GAR U-8-008 (Mule Deer)."""

    SIC_LEVEL = "Snow Interception Cover"
    MODERATE_AGE_LEVEL = "Moderate Zone Age > 20 years"

    MODERATE_AGE_TARGET = 0.67
    MAX_OPEN_ROAD_DENSITY = 3.0

    SHALLOW = "Shallow"
    MODERATE = "Moderate"
    DEEP = "Deep"

    # Schedule 1, Table 1. Values are proportions of each planning cell's current net area.
    TABLE1_TARGET_PERCENT = {
        1: 0.200, 2: 0.187, 3: 0.151, 4: 0.195, 5: 0.200, 6: 0.157, 7: 0.195, 8: 0.168,
        9: 0.209, 10: 0.173, 11: 0.199, 12: 0.200, 13: 0.199, 14: 0.200, 15: 0.200,
        16: 0.200, 17: 0.160, 18: 0.270, 19: 0.285, 20: 0.210, 21: 0.213, 22: 0.200,
        23: 0.199, 24: 0.198, 25: 0.193, 26: 0.190, 27: 0.176, 28: 0.200, 29: 0.200,
        30: 0.164, 31: 0.200, 32: 0.200, 33: 0.192, 34: 0.177, 35: 0.200, 36: 0.200,
        37: 0.200, 38: 0.200, 39: 0.200, 40: 0.200, 41: 0.180, 42: 0.200, 43: 0.182,
        44: 0.183, 46: 0.200, 47: 0.215, 48: 0.200, 49: 0.179, 50: 0.173, 51: 0.400,
        52: 0.199, 53: 0.150, 54: 0.180, 55: 0.290, 56: 0.246, 57: 0.229, 58: 0.254,
        59: 0.249, 60: 0.218, 61: 0.185, 62: 0.226, 63: 0.223, 64: 0.203, 65: 0.214,
        66: 0.162, 67: 0.268, 68: 0.217, 69: 0.277, 70: 0.186, 71: 0.325, 72: 0.250,
        73: 0.223, 74: 0.211, 77: 0.260, 78: 0.250, 79: 0.250, 80: 0.198, 81: 0.200,
        82: 0.400, 83: 0.400, 84: 0.250, 85: 0.250, 86: 0.250, 87: 0.200, 88: 0.200,
    }

    FMER_EXCLUSIONS = frozenset({"OPEN_FOREST", "OPEN_RANGE"})

    def __init__(self, gar, output_xls, logger, gar_config):
        self.gar = gar
        self.output_xls = output_xls
        self.logger = logger
        self.gar_config = gar_config

        self.dict_total_area = defaultdict(TotalArea)
        self.dict_cell_area = defaultdict(CellArea)
        self.dict_target = defaultdict(float)
        self.dict_zero_target = defaultdict(set)
        self.dict_cell_stats = defaultdict(U8008CellStats)
        self.lst_cells = []

        self.lst_level = [self.SIC_LEVEL, self.MODERATE_AGE_LEVEL]
        self.lst_headers = [
            "Planning Cell",
            "Gross Area Processed (ha)",
            "Provisional Net Area (ha)",
            "Table 1 SIC Target",
            "Calculated SIC Target (ha)",
            "Qualifying SIC (ha)",
            "SIC +/- (ha)",
            "SIC Assessment",
            "Moderate Zone Area (ha)",
            "Moderate Zone Age >20 (ha)",
            "Moderate Zone Age >20 (minimum 67%)",
            "Moderate-Age Assessment",
            "Open Road Density (maximum 3 km/km2)",
            "Access Assessment",
            "Data Completeness",
            "Spatial Assessment",
        ]
        self.lst_footers = [
            "Table 1 SIC targets are applied to the current net area using each legal planning-cell percentage.",
            "Net area excludes FMER Open Forest and Open Range. Results are provisional until FMER classes are supplied.",
            "IDFdm1 requires elevation and aspect to distinguish Shallow from Moderate Snowpack Zone.",
            "Moderate-zone age targets may be achieved across no more than three adjacent planning cells; adjacency must be verified externally.",
            "Road density remains Not assessed until accessible open-road lengths are supplied.",
            "Road/trail development in required SIC and other operational access measures require proposal-level review.",
            "U-8-008 does not apply to Boundary TSA woodlots. U-8-008 harvesting measures take precedence over U-8-007 overlaps.",
        ]

    @staticmethod
    def _number(value):
        if value is None or value == "":
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @classmethod
    def _area(cls, value):
        area = cls._number(value)
        return area if area is not None and area >= 0 else None

    @staticmethod
    def _normalize_bec(value):
        return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())

    @staticmethod
    def _cell_number(pcell):
        """Extract the Table 1 cell number from numeric or formatted UWR cell identifiers."""
        if isinstance(pcell, int):
            return pcell
        if isinstance(pcell, float) and pcell.is_integer():
            return int(pcell)
        matches = re.findall(r"\d+", str(pcell or ""))
        if not matches:
            return None
        # Formatted values such as U-8-008-12 should resolve to cell 12.
        return int(matches[-1])

    @classmethod
    def _snowpack_zone(cls, bec, elevation=None, aspect=None, override=None):
        if override:
            normalized = str(override).strip().title()
            if normalized in {cls.SHALLOW, cls.MODERATE, cls.DEEP}:
                return normalized

        bec_code = cls._normalize_bec(bec)
        if bec_code.startswith(("PPXH", "IDFXH")):
            return cls.SHALLOW
        if bec_code.startswith(("ICHDW", "MS")):
            return cls.MODERATE
        if bec_code.startswith(("ICHMK1", "ICHMW2", "ESSF")):
            return cls.DEEP

        if bec_code.startswith("IDFDM1"):
            elevation_value = cls._number(elevation)
            aspect_value = cls._number(aspect)
            if elevation_value is None or aspect_value is None:
                return None
            if elevation_value < 1000 and 135 <= aspect_value <= 270:
                return cls.SHALLOW
            return cls.MODERATE

        return None

    @classmethod
    def _qualifies_as_sic(cls, zone, bec, age, crown_closure, species):
        age_value = cls._number(age)
        closure_value = cls._number(crown_closure)
        species_code = str(species or "").strip().upper()

        # Existing GAR classes interpret Douglas-fir-leading VRI codes using startswith('F').
        if not species_code.startswith("F") or age_value is None or closure_value is None:
            return False

        bec_code = cls._normalize_bec(bec)
        if zone == cls.SHALLOW:
            return age_value >= 101 and closure_value >= 16
        if zone == cls.MODERATE:
            minimum_age = 121 if bec_code.startswith("ICHDW") else 101
            return age_value >= minimum_age and closure_value >= 45
        if zone == cls.DEEP:
            return age_value >= 121 and closure_value >= 55
        return False

    def calculate_level(
        self,
        bec,
        age,
        spp,
        cc,
        slp,
        thlb,
        diam,
        pct,
        gfa,
        notes,
        op_area,
        pcell,
        shp_area,
        target,
        height,
        fmer_class=None,
        elevation=None,
        aspect=None,
        snowpack_zone=None,
        is_woodlot=False,
        is_park=False,
    ):
        """Classify a resultant polygon and accumulate U-8-008 planning-cell statistics."""
        if is_woodlot or is_park:
            return None

        area = self._area(shp_area)
        if area is None or area == 0:
            return None

        stats = self.dict_cell_stats[pcell]
        stats.gross_hectares += area

        fmer_value = str(fmer_class or "").strip().upper().replace(" ", "_")
        if fmer_value in self.FMER_EXCLUSIONS:
            return None
        if not fmer_value:
            stats.fmer_unassessed_hectares += area

        stats.net_hectares += area
        op_area = op_area or ""
        cell = self.dict_cell_area[pcell]
        op_cell = self.dict_total_area[op_area].pcell[pcell]
        cell.hectares += area
        op_cell.hectares += area

        zone = self._snowpack_zone(
            bec=bec,
            elevation=elevation,
            aspect=aspect,
            override=snowpack_zone,
        )
        levels = []

        if zone is None:
            stats.unresolved_snowpack_hectares += area
        else:
            if self._qualifies_as_sic(zone, bec, age, cc, spp):
                levels.append(self.SIC_LEVEL)
                stats.sic_hectares += area
                cell.level[self.SIC_LEVEL].hectares += area
                op_cell.level[self.SIC_LEVEL].hectares += area

            if zone == self.MODERATE:
                stats.moderate_hectares += area
                age_value = self._number(age)
                if age_value is not None and age_value > 20:
                    levels.append(self.MODERATE_AGE_LEVEL)
                    stats.moderate_age_over_20_hectares += area
                    cell.level[self.MODERATE_AGE_LEVEL].hectares += area
                    op_cell.level[self.MODERATE_AGE_LEVEL].hectares += area

        if pcell not in self.lst_cells:
            self.lst_cells.append(pcell)
        if op_area and target == 0:
            self.dict_zero_target[op_area].add(pcell)

        return "; ".join(levels) if levels else None

    def record_open_road_length(self, pcell, length_km, developed_or_proposed_in_sic=False):
        """
        Record accessible open-road length for one planning cell.

        Supply zero kilometres to explicitly mark a cell as assessed with no open roads. Input lines must be clipped
        and de-duplicated by planning cell before this method is called.
        """
        length = self._number(length_km)
        if length is None or length < 0:
            return False

        stats = self.dict_cell_stats[pcell]
        stats.road_density_assessed = True
        stats.open_road_km += length
        if developed_or_proposed_in_sic:
            stats.access_in_sic_conflict = True
        return True

    def calculate_targets(self):
        """Calculate current target hectares from Table 1 percentages and accumulated net area."""
        self.logger.info("Calculating U-8-008 planning-cell targets")

        for pcell, cell in self.dict_cell_area.items():
            cell_number = self._cell_number(pcell)
            target_percent = self.TABLE1_TARGET_PERCENT.get(cell_number)
            if target_percent is not None:
                target_hectares = cell.hectares * target_percent
                cell.target = target_hectares
                cell.level[self.SIC_LEVEL].target = target_hectares
                self.dict_target[pcell] = target_hectares
            cell.level[self.MODERATE_AGE_LEVEL].target = (
                self.dict_cell_stats[pcell].moderate_hectares * self.MODERATE_AGE_TARGET
            )

    @staticmethod
    def _percent(numerator, denominator):
        return numerator / denominator if denominator > 0 else None

    def evaluate_moderate_age_group(self, pcells):
        """Evaluate the 67% Moderate Snowpack Zone age measure across one to three cells."""
        cells = list(dict.fromkeys(pcells))
        if not 1 <= len(cells) <= 3:
            raise ValueError("The U-8-008 moderate-age measure may use one to three planning cells.")

        total = sum(self.dict_cell_stats[cell].moderate_hectares for cell in cells)
        qualifying = sum(self.dict_cell_stats[cell].moderate_age_over_20_hectares for cell in cells)
        percent = self._percent(qualifying, total)
        return {
            "pcells": cells,
            "moderate_hectares": total,
            "age_over_20_hectares": qualifying,
            "percent": percent,
            "status": "Not applicable" if percent is None else ("Meets" if percent >= self.MODERATE_AGE_TARGET else "Deficit"),
            "adjacency_verified": False,
        }

    def evaluate_cell(self, pcell):
        """Return the available U-8-008 spatial metrics for one planning cell."""
        stats = self.dict_cell_stats[pcell]
        cell_number = self._cell_number(pcell)
        target_percent = self.TABLE1_TARGET_PERCENT.get(cell_number)
        target_hectares = stats.net_hectares * target_percent if target_percent is not None else None

        if target_hectares is None:
            sic_difference = None
            sic_numeric_status = "Target unavailable"
        else:
            sic_difference = stats.sic_hectares - target_hectares
            sic_numeric_status = "Meets" if sic_difference >= 0 else "Deficit"

        moderate_percent = self._percent(
            stats.moderate_age_over_20_hectares,
            stats.moderate_hectares,
        )
        if moderate_percent is None:
            moderate_status = "Not applicable"
        elif moderate_percent >= self.MODERATE_AGE_TARGET:
            moderate_status = "Meets in cell"
        else:
            moderate_status = "Below 67% in cell - grouping may apply"

        if stats.road_density_assessed and stats.gross_hectares > 0:
            road_density = stats.open_road_km / (stats.gross_hectares / 100.0)
            access_status = "Meets" if road_density <= self.MAX_OPEN_ROAD_DENSITY else "Deficit"
            if stats.access_in_sic_conflict:
                access_status = "Conflict - access recorded in required SIC"
        else:
            road_density = None
            access_status = "Not assessed"

        missing = []
        if stats.fmer_unassessed_hectares > 0:
            missing.append("FMER")
        if stats.unresolved_snowpack_hectares > 0:
            missing.append("snowpack")
        if not stats.road_density_assessed:
            missing.append("road density")
        completeness = "Complete inputs" if not missing else "Missing: {}".format(", ".join(missing))

        has_deficit = sic_numeric_status == "Deficit" or access_status in {
            "Deficit",
            "Conflict - access recorded in required SIC",
        }
        if has_deficit:
            assessment = "Deficit or access conflict in assessed measures"
        elif missing:
            assessment = "Incomplete - one or more required inputs not assessed"
        elif moderate_status.startswith("Below"):
            assessment = "Incomplete - moderate-zone cell grouping must be assessed"
        elif sic_numeric_status == "Meets" and access_status == "Meets":
            assessment = "Meets assessed spatial thresholds"
        else:
            assessment = "Incomplete spatial assessment"

        return {
            "pcell": pcell,
            "gross_hectares": stats.gross_hectares,
            "net_hectares": stats.net_hectares,
            "target_percent": target_percent,
            "target_hectares": target_hectares,
            "sic_hectares": stats.sic_hectares,
            "sic_difference": sic_difference,
            "sic_status": sic_numeric_status,
            "moderate_hectares": stats.moderate_hectares,
            "moderate_age_over_20_hectares": stats.moderate_age_over_20_hectares,
            "moderate_percent": moderate_percent,
            "moderate_status": moderate_status,
            "road_density": road_density,
            "access_status": access_status,
            "completeness": completeness,
            "assessment": assessment,
        }

    def write_excel(self):
        """Write the available U-8-008 planning-cell assessment to Excel."""
        self.logger.info("Writing U-8-008 results to Excel")
        self.calculate_targets()

        workbook = xlsxwriter.Workbook(filename=self.output_xls)
        gar_excel = GARExcel(wb=workbook)
        worksheet = workbook.add_worksheet(name="U-8-008")

        date_now = dt.today().strftime("%B, %Y")
        worksheet.write(0, 0, "Created: {}. GAR ORDER: {}".format(date_now, self.gar))

        for column, header in enumerate(self.lst_headers):
            worksheet.write(1, column, header, gar_excel.black_style_bottom_border)

        row = 2
        for pcell in sorted(self.dict_cell_stats, key=lambda value: str(value)):
            result = self.evaluate_cell(pcell)
            values = [
                result["pcell"],
                result["gross_hectares"],
                result["net_hectares"],
                result["target_percent"],
                result["target_hectares"],
                result["sic_hectares"],
                result["sic_difference"],
                result["sic_status"],
                result["moderate_hectares"],
                result["moderate_age_over_20_hectares"],
                result["moderate_percent"],
                result["moderate_status"],
                result["road_density"],
                result["access_status"],
                result["completeness"],
                result["assessment"],
            ]

            for column, value in enumerate(values):
                if isinstance(value, float) and column not in (3, 10):
                    value = gar_excel.round_value(value)

                if column in (3, 10):
                    style = gar_excel.black_percent_style
                    if column == 10 and value is not None and value < self.MODERATE_AGE_TARGET:
                        style = gar_excel.red_letters_percent
                elif column == 6 and value is not None and value < 0:
                    style = gar_excel.red_letters
                elif column in (7, 13, 15) and isinstance(value, str) and (
                    value.startswith("Deficit") or value.startswith("Conflict")
                ):
                    style = gar_excel.red_letters
                else:
                    style = gar_excel.black_style

                worksheet.write(row, column, value, style)
            row += 1

        row += 1
        for footer in self.lst_footers:
            worksheet.merge_range(row, 0, row, len(self.lst_headers) - 1, footer, gar_excel.black_style)
            row += 1

        worksheet.freeze_panes(2, 1)
        worksheet.set_column(0, 0, 18)
        worksheet.set_column(1, 14, 20)
        worksheet.set_column(15, 15, 46)
        workbook.close()

