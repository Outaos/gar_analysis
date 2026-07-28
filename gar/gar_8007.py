"""
----------------------------------------------------------------------------------------------------------------
    PYTHON SCRIPT: gar_8007.py

    Purpose:      Class containing the spatial assessment rules for GAR U-8-007 - Moose.
    Source:       https://www.env.gov.bc.ca/wld/documents/uwr/U-8-007_ord.pdf

    The legal Schedule 1 thresholds implemented here are:
      * at least 20% of each planning cell in stands more than 16 metres high;
      * at least 60% of each planning cell in stands more than 30 years old; and
      * at least 50% of the eligible riparian management area in stands more than 16 metres high.

    Riparian management areas are limited to S1, S2, S3 and S5 streams and W1, W3 and W5 wetlands.
    The general GAR analysis cursor does not supply riparian attributes, so callers must populate that measure
    through record_riparian_area(). Until they do, the report marks the riparian measure as not assessed.
----------------------------------------------------------------------------------------------------------------
"""

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime as dt

import xlsxwriter

from util.gar_classes import GARExcel, TotalArea, CellArea


@dataclass
class U8007CellStats:
    """Areas used to assess the three spatial U-8-007 measures for one planning cell."""

    total_hectares: float = 0.0
    height_over_16_hectares: float = 0.0
    age_over_30_hectares: float = 0.0
    riparian_total_hectares: float = 0.0
    riparian_height_over_16_hectares: float = 0.0
    riparian_assessed: bool = False


class Gar8007:
    """Spatial assessment and Excel reporting for GAR U-8-007 (Moose)."""

    HEIGHT_LEVEL = "Stand Height > 16 m"
    AGE_LEVEL = "Stand Age > 30 years"

    HEIGHT_TARGET = 0.20
    AGE_TARGET = 0.60
    RIPARIAN_HEIGHT_TARGET = 0.50

    ELIGIBLE_STREAM_CLASSES = frozenset({"S1", "S2", "S3", "S5"})
    ELIGIBLE_WETLAND_CLASSES = frozenset({"W1", "W3", "W5"})

    def __init__(self, gar, output_xls, logger, gar_config):
        self.gar = gar
        self.output_xls = output_xls
        self.logger = logger
        self.gar_config = gar_config

        # These objects preserve the interface used by the existing GAR analysis framework.
        self.dict_total_area = defaultdict(TotalArea)
        self.dict_cell_area = defaultdict(CellArea)
        self.dict_target = defaultdict(float)
        self.dict_zero_target = defaultdict(set)
        self.lst_cells = []

        # U-8-007-specific statistics, including the riparian measure that CellArea cannot represent directly.
        self.dict_cell_stats = defaultdict(U8007CellStats)

        self.lst_level = [self.HEIGHT_LEVEL, self.AGE_LEVEL]
        self.lst_headers = [
            "Planning Cell",
            "Net Assessed Planning Cell Area (ha)",
            "Height >16 m Target (ha)",
            "Height >16 m Area (ha)",
            "Height >16 m (minimum 20%)",
            "Age >30 Target (ha)",
            "Age >30 Area (ha)",
            "Age >30 (minimum 60%)",
            "Eligible RMA Area (ha)",
            "RMA Height >16 m Target (ha)",
            "RMA Height >16 m Area (ha)",
            "RMA Height >16 m (minimum 50%)",
            "Spatial Assessment",
        ]
        self.lst_footers = [
            "Legal thresholds use 'more than 16 metres' and 'more than 30 years'; comparisons are strict (>).",
            "Eligible riparian management areas are S1, S2, S3 and S5 streams and W1, W3 and W5 wetlands.",
            "Riparian results remain Not assessed until record_riparian_area() is supplied with non-overlapping RMA polygons.",
            "U-8-007 does not apply to woodlots in the Boundary TSA or where U-8-007 overlaps U-8-008.",
            "The report is a spatial screening result; salvage, subsurface-resource and exemption provisions require external review.",
        ]

    @staticmethod
    def _number(value):
        """Return a float for usable numeric values, otherwise None."""
        if value is None or value == "":
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _area(value):
        """Normalize an area value and reject null, negative and non-numeric values."""
        area = Gar8007._number(value)
        if area is None or area < 0:
            return None
        return area

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
        is_woodlot=False,
        overlaps_u8008=False,
    ):
        """
        Classify a resultant polygon and accumulate planning-cell statistics.

        The existing GAR framework supplies the parameters through height. The two optional flags support the order's
        spatial exclusions when a future caller has those overlays available. They default to False so the method remains
        compatible with the current generic cursor call.

        The denominator is the planning-cell area presented to this class. Polygons without usable height or age still
        contribute to that denominator but do not contribute to the qualifying numerator.
        """
        if is_woodlot or overlaps_u8008:
            return None

        area = self._area(shp_area)
        if area is None or area == 0:
            return None

        op_area = op_area or ""
        cell = self.dict_cell_area[pcell]
        op_cell = self.dict_total_area[op_area].pcell[pcell]
        stats = self.dict_cell_stats[pcell]

        cell.hectares += area
        op_cell.hectares += area
        stats.total_hectares += area

        levels = []
        height_value = self._number(height)
        age_value = self._number(age)

        # The legal Schedule uses "more than", so values exactly equal to 16 m or 30 years do not qualify.
        if height_value is not None and height_value > 16:
            levels.append(self.HEIGHT_LEVEL)
            cell.level[self.HEIGHT_LEVEL].hectares += area
            op_cell.level[self.HEIGHT_LEVEL].hectares += area
            stats.height_over_16_hectares += area

        if age_value is not None and age_value > 30:
            levels.append(self.AGE_LEVEL)
            cell.level[self.AGE_LEVEL].hectares += area
            op_cell.level[self.AGE_LEVEL].hectares += area
            stats.age_over_30_hectares += area

        if pcell not in self.lst_cells:
            self.lst_cells.append(pcell)

        if op_area and target == 0:
            self.dict_zero_target[op_area].add(pcell)

        return "; ".join(levels) if levels else None

    def record_riparian_area(
        self,
        pcell,
        shp_area,
        height,
        classification,
        is_woodlot=False,
        overlaps_u8008=False,
    ):
        """
        Accumulate a non-overlapping eligible riparian-management-area polygon.

        Args:
            pcell: U-8-007 planning-cell identifier.
            shp_area: Polygon area in hectares.
            height: Stand height in metres.
            classification: Stream/wetland class; one of S1, S2, S3, S5, W1, W3 or W5.
            is_woodlot: True when the polygon is in an excluded Boundary TSA woodlot.
            overlaps_u8008: True where U-8-008 takes precedence.

        Returns:
            bool: True when the record was an eligible RMA polygon and was accumulated.

        Callers must erase/dissolve overlapping RMA polygons before calling this method so area is not double counted.
        """
        if is_woodlot or overlaps_u8008:
            return False

        rma_class = str(classification or "").strip().upper()
        eligible_classes = self.ELIGIBLE_STREAM_CLASSES | self.ELIGIBLE_WETLAND_CLASSES
        if rma_class not in eligible_classes:
            return False

        area = self._area(shp_area)
        if area is None:
            return False

        stats = self.dict_cell_stats[pcell]
        stats.riparian_assessed = True
        stats.riparian_total_hectares += area

        height_value = self._number(height)
        if height_value is not None and height_value > 16:
            stats.riparian_height_over_16_hectares += area

        return True

    def calculate_targets(self):
        """Calculate target hectares for the cell-wide height and age measures."""
        self.logger.info("Calculating U-8-007 planning-cell targets")

        for pcell, cell in self.dict_cell_area.items():
            height_target = cell.hectares * self.HEIGHT_TARGET
            age_target = cell.hectares * self.AGE_TARGET
            cell.level[self.HEIGHT_LEVEL].target = height_target
            cell.level[self.AGE_LEVEL].target = age_target
            self.dict_target[pcell] = {
                self.HEIGHT_LEVEL: height_target,
                self.AGE_LEVEL: age_target,
            }

        # Operating-area subdivisions are retained for framework compatibility, but legal compliance is assessed by
        # complete planning cell in evaluate_cell() and write_excel().
        for total_area in self.dict_total_area.values():
            for cell in total_area.pcell.values():
                cell.level[self.HEIGHT_LEVEL].target = cell.hectares * self.HEIGHT_TARGET
                cell.level[self.AGE_LEVEL].target = cell.hectares * self.AGE_TARGET

    @staticmethod
    def _percent(numerator, denominator):
        return numerator / denominator if denominator > 0 else None

    @staticmethod
    def _measure_status(percent, target, unavailable="Not assessed"):
        if percent is None:
            return unavailable
        return "Meets" if percent >= target else "Deficit"

    def evaluate_cell(self, pcell):
        """Return all U-8-007 spatial metrics for one planning cell."""
        stats = self.dict_cell_stats[pcell]

        height_percent = self._percent(stats.height_over_16_hectares, stats.total_hectares)
        age_percent = self._percent(stats.age_over_30_hectares, stats.total_hectares)
        height_status = self._measure_status(height_percent, self.HEIGHT_TARGET, "No cell area")
        age_status = self._measure_status(age_percent, self.AGE_TARGET, "No cell area")

        if not stats.riparian_assessed:
            riparian_percent = None
            riparian_status = "Not assessed"
        elif stats.riparian_total_hectares == 0:
            riparian_percent = None
            riparian_status = "No eligible RMA"
        else:
            riparian_percent = self._percent(
                stats.riparian_height_over_16_hectares,
                stats.riparian_total_hectares,
            )
            riparian_status = self._measure_status(riparian_percent, self.RIPARIAN_HEIGHT_TARGET)

        statuses = (height_status, age_status, riparian_status)
        if "Deficit" in statuses:
            assessment = "Deficit in one or more spatial measures"
        elif riparian_status == "Not assessed":
            assessment = "Incomplete - riparian measure not assessed"
        elif height_status == "Meets" and age_status == "Meets" and riparian_status in {"Meets", "No eligible RMA"}:
            assessment = "Meets assessed spatial thresholds"
        else:
            assessment = "Incomplete spatial assessment"

        return {
            "pcell": pcell,
            "total_hectares": stats.total_hectares,
            "height_target_hectares": stats.total_hectares * self.HEIGHT_TARGET,
            "height_hectares": stats.height_over_16_hectares,
            "height_percent": height_percent,
            "height_status": height_status,
            "age_target_hectares": stats.total_hectares * self.AGE_TARGET,
            "age_hectares": stats.age_over_30_hectares,
            "age_percent": age_percent,
            "age_status": age_status,
            "riparian_total_hectares": stats.riparian_total_hectares,
            "riparian_target_hectares": stats.riparian_total_hectares * self.RIPARIAN_HEIGHT_TARGET,
            "riparian_height_hectares": stats.riparian_height_over_16_hectares,
            "riparian_percent": riparian_percent,
            "riparian_status": riparian_status,
            "assessment": assessment,
        }

    def write_excel(self):
        """Write one legal planning-cell summary sheet to the configured Excel workbook."""
        self.logger.info("Writing U-8-007 results to Excel")
        self.calculate_targets()

        workbook = xlsxwriter.Workbook(filename=self.output_xls)
        gar_excel = GARExcel(wb=workbook)
        worksheet = workbook.add_worksheet(name="U-8-007")

        date_now = dt.today().strftime("%B, %Y")
        worksheet.write(0, 0, "Created: {}. GAR ORDER: {}".format(date_now, self.gar))

        for column, header in enumerate(self.lst_headers):
            worksheet.write(1, column, header, gar_excel.black_style_bottom_border)

        row = 2
        for pcell in sorted(self.dict_cell_stats, key=lambda value: str(value)):
            result = self.evaluate_cell(pcell)
            values = [
                result["pcell"],
                result["total_hectares"],
                result["height_target_hectares"],
                result["height_hectares"],
                result["height_percent"],
                result["age_target_hectares"],
                result["age_hectares"],
                result["age_percent"],
                result["riparian_total_hectares"] if result["riparian_status"] != "Not assessed" else None,
                result["riparian_target_hectares"] if result["riparian_status"] != "Not assessed" else None,
                result["riparian_height_hectares"] if result["riparian_status"] != "Not assessed" else None,
                result["riparian_percent"],
                result["assessment"],
            ]

            for column, value in enumerate(values):
                if isinstance(value, float) and column not in (4, 7, 11):
                    value = gar_excel.round_value(value)

                if column == 4:
                    style = (
                        gar_excel.red_letters_percent
                        if value is not None and value < self.HEIGHT_TARGET
                        else gar_excel.black_percent_style
                    )
                elif column == 7:
                    style = (
                        gar_excel.red_letters_percent
                        if value is not None and value < self.AGE_TARGET
                        else gar_excel.black_percent_style
                    )
                elif column == 11:
                    style = (
                        gar_excel.red_letters_percent
                        if value is not None and value < self.RIPARIAN_HEIGHT_TARGET
                        else gar_excel.black_percent_style
                    )
                elif column == 12 and result["assessment"].startswith("Deficit"):
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
        worksheet.set_column(1, 11, 19)
        worksheet.set_column(12, 12, 42)
        workbook.close()

