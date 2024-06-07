using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;        // for Path class: https://learn.microsoft.com/ja-jp/dotnet/api/system.io.path?view=net-7.0
using System.Windows;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Security.Cryptography;
using System.Windows.Data;

namespace constraints_summarizer_esapi_v15_5
{
    internal class HeaderInfo
    {
        public string PlanName {get;set;}
        public string Energy {get;set;}
        public string Machine {get;set;}
        public string PatientName {get;set;}
        public string PatientId {get;set;}
//        public string Oncologist {get;set;}
//        public string Planner {get;set;}
//        public string PrescribedDose {get;set;}
//        public string Normalization {get;set;}
        public string CalcAlgorithm {get;set;}
        public string CalcGridSize {get;set;}
        public string PlanNormalizationMethod {get;set;}
    }
    internal class ConstraintsFormGen
    {
        /* parameters */
        // Header
        const string TITLE_CELL_POS = "B2";
        const Int32 LABEL_WID = 2;
        const Int32 VALUE_WID = 4;
        const string PLAN_NAME_LABEL_POS = "B4";
        const string PATIENT_NAME_LABEL_POS = "B5";
        const string PATIENT_ID_LABEL_POS = "B6";
        const string ONCOLOGIST_LABEL_POS = "H5";
        const string PLANNER_LABEL_POS = "H6";
        const string DOSE_PER_FRACTION_LABEL_POS = "B7";
        const string PRESCRIPTION_LABEL_POS = "H7";
        const string ENERG_LABEL_POS = "B8";
        const string MACHINE_LABEL_POS = "H8";
        const string CALC_ALGORITHM_LABEL_POS = "B9";
        const string CALC_GRID_SIZE_LABEL_POS = "H9";
        const string PLAN_NORMALIZATION_METHOD_LABEL_POS = "B10";

        // Body
        const Int32 STRUCTURE_NAME_COL = 2;
        const Int32 DVH_INDEX_COL = STRUCTURE_NAME_COL + 1;
        const Int32 RELATION_COL = STRUCTURE_NAME_COL + 2;
        const Int32 TOLERANCE_INDEX_VALUE_COL = STRUCTURE_NAME_COL + 3;
        const Int32 TOLERANCE_UNIT_COL = STRUCTURE_NAME_COL + 4;
        const Int32 ACCEPTABLE_INDEX_VALUE_COL = STRUCTURE_NAME_COL + 5;
        const Int32 ACCEPTABLE_UNIT_COL = STRUCTURE_NAME_COL + 6;
        const Int32 ACTUAL_DVH_INDEX_COL = STRUCTURE_NAME_COL + 7;
        const Int32 EQUAL_COL = STRUCTURE_NAME_COL + 8;
        const Int32 ACTUAL_INDEX_VALUE_COL = STRUCTURE_NAME_COL + 9;
        const Int32 ACTUAL_UNIT_COL = STRUCTURE_NAME_COL + 10;
        const Int32 DECISION_COL = STRUCTURE_NAME_COL + 11;

        // width
        const Int32 STRUCTURE_COL_WID = 15;
        const Int32 INDEX_COL_WID = 10;
        const Int32 RELATION_COL_WID = 4;
        const Int32 VALUE_AND_DECISION_COL_WID = 8;
        const Int32 UNIT_COL_WID = 6;

        // height
        const Int32 HEADER_HEIGHT = 22;
        const Int32 CI_ROW_HEIGHT = 33;
        const Int32 CONSTS_HEIGHT = 18;

        // sheet
        const string SHEET_NAME = "Summary";
        const string FONT_NAME = "TimesNewRoman";
        const double TITLE_FONT_SIZE = 26;
        const double HEADER_FONT_SIZE = 14;
        const double CONTENTS_FONT_SIZE = 11;
        const double DECISION_FONT_SIZE = 11;

        /* fields */
        private SLDocument sl;
        private ConstraintsReferenceParser consts_parser;
        private string output_file_path = null;
        private (string, DoseConstraints.ForCICalc)? for_ci = null;
        private HeaderInfo header_info;

        // style
        SLStyle header_style;
        SLStyle header_style_index;
        SLStyle header_style_value;
        SLStyle contents_style;
        SLStyle full_style;
        SLStyle start_style;
        SLStyle mid_style;
        SLStyle end_style;
        BorderStyleValues border_style = BorderStyleValues.Thin;

        /* properties */

        public ConstraintsFormGen(in ConstraintsReferenceParser parser, HeaderInfo header, string output_file)
        {

            if (parser.Constraints.Count == 0)
            {
                throw new ArgumentException("ERROR: Constraints parser does not have any parsed field.");
            }
            else
            {; }

            consts_parser = parser;
            header_info = header;

            /* For conformity index */
            DoseConstraints? null_or_ci = consts_parser.Constraints.FirstOrDefault(x => x.Index == DoseConstraints.DvhIndex.CI); // FirstOrDefault() returns null if matched value is not found.
            for_ci = (null_or_ci == null) ? null : (null_or_ci.Structure, null_or_ci.ForCI);

            /* Prepare for writing constraints file */
            output_file_path = Path.GetFullPath(output_file);

            // FileShare is not needed for shared(legacy) Excel file 
            sl = new SLDocument(output_file_path);


            header_style = sl.CreateStyle();
            header_style.Font.FontSize = HEADER_FONT_SIZE;
            header_style.Font.FontName = FONT_NAME;

            header_style_index = header_style.Clone();
            header_style_index.Border.TopBorder.BorderStyle = border_style;
            header_style_index.Border.BottomBorder.BorderStyle = border_style;
            header_style_index.Border.LeftBorder.BorderStyle = border_style;
            header_style_index.Border.RightBorder.BorderStyle = BorderStyleValues.None;
            header_style_index.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            header_style_index.Alignment.Vertical = VerticalAlignmentValues.Center;
            header_style_value = header_style.Clone();
            header_style_value.Border.TopBorder.BorderStyle = border_style;
            header_style_value.Border.BottomBorder.BorderStyle = border_style;
            header_style_value.Border.LeftBorder.BorderStyle = BorderStyleValues.None;
            header_style_value.Border.RightBorder.BorderStyle = border_style;
            header_style_value.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            header_style_value.Alignment.Vertical = VerticalAlignmentValues.Center;
            header_style_value.SetWrapText(true);

            contents_style = sl.CreateStyle();
            contents_style.Font.FontName = FONT_NAME;
            contents_style.SetWrapText(true);
            contents_style.Border.TopBorder.BorderStyle = border_style;
            contents_style.Border.BottomBorder.BorderStyle = border_style;
            contents_style.Border.LeftBorder.BorderStyle = border_style;
            contents_style.Border.RightBorder.BorderStyle = border_style;
            contents_style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            contents_style.Alignment.Vertical = VerticalAlignmentValues.Center;

            full_style = sl.CreateStyle();
            full_style.Font.FontName = FONT_NAME;
            full_style.SetWrapText(true);
            full_style.Border.TopBorder.BorderStyle = border_style;
            full_style.Border.BottomBorder.BorderStyle = border_style;
            full_style.Border.LeftBorder.BorderStyle = border_style;
            full_style.Border.RightBorder.BorderStyle = border_style;
            full_style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            full_style.Alignment.Vertical = VerticalAlignmentValues.Center;

            start_style = sl.CreateStyle();
            start_style.Font.FontName = FONT_NAME;
            start_style.Border.TopBorder.BorderStyle = border_style;
            start_style.Border.BottomBorder.BorderStyle = border_style;
            start_style.Border.LeftBorder.BorderStyle = border_style;
            start_style.Border.RightBorder.BorderStyle = BorderStyleValues.None;
            start_style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            start_style.Alignment.Vertical = VerticalAlignmentValues.Center;

            mid_style = sl.CreateStyle();
            mid_style.Font.FontName = FONT_NAME;
            mid_style.Border.TopBorder.BorderStyle = border_style;
            mid_style.Border.BottomBorder.BorderStyle = border_style;
            mid_style.Border.LeftBorder.BorderStyle = BorderStyleValues.None;
            mid_style.Border.RightBorder.BorderStyle = BorderStyleValues.None;
            mid_style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            mid_style.Alignment.Vertical = VerticalAlignmentValues.Center;

            end_style = sl.CreateStyle();
            end_style.Font.FontName = FONT_NAME;
            end_style.Border.TopBorder.BorderStyle = border_style;
            end_style.Border.BottomBorder.BorderStyle = border_style;
            end_style.Border.LeftBorder.BorderStyle = BorderStyleValues.None;
            end_style.Border.RightBorder.BorderStyle = border_style;
            end_style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            end_style.Alignment.Vertical = VerticalAlignmentValues.Center;

            SLPageSettings ps = new SLPageSettings();
            ps.TopMargin = 0.6;
            ps.BottomMargin = 0.6;
            ps.RightMargin = 0.6;
            ps.LeftMargin = 0.6;
            ps.PrintHorizontalCentered = true;
            ps.PrintVerticalCentered = false;
            sl.SetPageSettings(ps);

        }

        public void WriteConstraintsSummarySheet()
        {
            sl.AddWorksheet(SHEET_NAME);
            sl.MoveWorksheet(SHEET_NAME, 1);
            sl.SelectWorksheet(SHEET_NAME);
            bool _ = sl.DeleteWorksheet(ConstraintsReferenceParser.ListSheetName);


            /* Set width of each column */
            sl.SetColumnWidth(STRUCTURE_NAME_COL, STRUCTURE_COL_WID);
            sl.SetColumnWidth(DVH_INDEX_COL, INDEX_COL_WID);
            sl.SetColumnWidth(RELATION_COL, RELATION_COL_WID);
            sl.SetColumnWidth(TOLERANCE_INDEX_VALUE_COL, VALUE_AND_DECISION_COL_WID);
            sl.SetColumnWidth(TOLERANCE_UNIT_COL, UNIT_COL_WID);
            sl.SetColumnWidth(ACCEPTABLE_INDEX_VALUE_COL, VALUE_AND_DECISION_COL_WID);
            sl.SetColumnWidth(ACCEPTABLE_UNIT_COL, UNIT_COL_WID);
            sl.SetColumnWidth(ACTUAL_DVH_INDEX_COL, INDEX_COL_WID);
            sl.SetColumnWidth(EQUAL_COL, UNIT_COL_WID);
            sl.SetColumnWidth(ACTUAL_INDEX_VALUE_COL, VALUE_AND_DECISION_COL_WID);
            sl.SetColumnWidth(ACTUAL_UNIT_COL, UNIT_COL_WID);
            sl.SetColumnWidth(DECISION_COL, VALUE_AND_DECISION_COL_WID);

            /* Header */
            (int, int) pos = SpreadSheetLightExtensions.GetRowColIndexFromCellReference(TITLE_CELL_POS);
            sl.SetCellValue(TITLE_CELL_POS, consts_parser.Title);
            sl.MergeWorksheetCells(pos.Item1, pos.Item2, pos.Item1, pos.Item2 + (2 * LABEL_WID) + (2 * VALUE_WID) - 1);
            SLStyle title_style = sl.CreateStyle();
            title_style.Font.FontSize = TITLE_FONT_SIZE;
            //            title_style.Font.FontName = FONT_NAME;
            title_style.Font.Bold = true;
            title_style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            title_style.Alignment.Vertical = VerticalAlignmentValues.Center;
            sl.SetCellStyle(TITLE_CELL_POS, title_style);

            WriteHeaderProperty(sl, PLAN_NAME_LABEL_POS, "Plan Name", header_info.PlanName, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, PATIENT_NAME_LABEL_POS, "Patient Name", header_info.PatientName, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, PATIENT_ID_LABEL_POS, "Patient ID", header_info.PatientId, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, ONCOLOGIST_LABEL_POS, "Oncologist", consts_parser.Oncologist, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, PLANNER_LABEL_POS, "Planner", consts_parser.Planner, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, DOSE_PER_FRACTION_LABEL_POS, "Prescribed Dose", string.Format("{0}Gy/{1}fr   ({2}Gy/fr)", consts_parser.TotalDose, consts_parser.Fraction, (consts_parser.TotalDose / consts_parser.Fraction)), LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, PRESCRIPTION_LABEL_POS, "Normalization", string.Format("{0:f1}% of the PD", consts_parser.PrescriptionPercentage), LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, ENERG_LABEL_POS, "Energy", header_info.Energy, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, MACHINE_LABEL_POS, "Machine", header_info.Machine, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, CALC_ALGORITHM_LABEL_POS, "Calc. Algorithm", header_info.CalcAlgorithm, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, CALC_GRID_SIZE_LABEL_POS, "Grid size (cm)", header_info.CalcGridSize, LABEL_WID, VALUE_WID);
            WriteHeaderProperty(sl, PLAN_NORMALIZATION_METHOD_LABEL_POS, "Plan normalization method", header_info.PlanNormalizationMethod, LABEL_WID * 2, /*VALUE_WID * 2*/6);

            Int32 row_num = SpreadSheetLightExtensions.GetRowColIndexFromCellReference(PLAN_NORMALIZATION_METHOD_LABEL_POS).Item1 + 2;

            /* CI if it's needed. */
            if (for_ci.HasValue)
            {
                string target_volume_name = for_ci.Value.Item1;
                DoseConstraints.ForCICalc ci = for_ci.Value.Item2;


                WriteCIRow(sl, ref row_num, "計画標的体積",
                   $"V_{{{target_volume_name}}}", ci.TargetVolumeCc);

                WriteCIRow(sl, ref row_num, "PTV内の処方線量が照射される体積",
                    $"V_{{{target_volume_name},ref}}", ci.PtvIrradiatedVolumeCc);

                WriteCIRow(sl, ref row_num, "全体積内の処方線量が照射される体積",
                    "V_{ref}", ci.AllIrradiatedVolumeCc);

            }
            else {; }


            /* Body of constraints */

            // contents
            contents_style.Fill.SetPatternType(PatternValues.Solid);
            contents_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xb7dee8));

            Int32 contents_row = row_num;
            sl.SetRowHeight(contents_row, 30);

            sl.SetCellValueAndStyle(contents_row, STRUCTURE_NAME_COL, "Structure name", contents_style);

            sl.SetCellValue(contents_row, DVH_INDEX_COL, "Criteria");
            sl.MergeWorksheetCells(contents_row, DVH_INDEX_COL, contents_row, TOLERANCE_UNIT_COL, contents_style);

            sl.SetCellValue(contents_row, ACCEPTABLE_INDEX_VALUE_COL, "Acceptable\ncriteria");
            sl.MergeWorksheetCells(contents_row, ACCEPTABLE_INDEX_VALUE_COL, contents_row, ACCEPTABLE_UNIT_COL, contents_style);


            contents_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xffc000));

            sl.SetCellValue(contents_row, ACTUAL_DVH_INDEX_COL, "Actual plan");
            sl.MergeWorksheetCells(contents_row, ACTUAL_DVH_INDEX_COL, contents_row, ACTUAL_UNIT_COL, contents_style);

            sl.SetCellValueAndStyle(contents_row, DECISION_COL, "Result", contents_style);

            ++row_num;

            foreach (var consts_pair in consts_parser.ConstraintsDictTarget.Dict)
            {
                WriteConstsRowsInSummary(sl, row_num, consts_pair.Key, consts_pair.Value);
                row_num += consts_pair.Value.Count;
            }

            mid_style.Alignment.Horizontal = HorizontalAlignmentValues.Left;
            sl.SetCellValueAndStyle(row_num, STRUCTURE_NAME_COL, "The OAR constraints are listed below.", mid_style);
            mid_style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
            ++row_num;

            foreach (var consts_pair in consts_parser.ConstraintsDictOAR.Dict)
            {
                WriteConstsRowsInSummary(sl, row_num, consts_pair.Key, consts_pair.Value);
                row_num += consts_pair.Value.Count;
            }

            /* Footer */
            ++row_num;
            sl.SetCellValue(row_num, STRUCTURE_NAME_COL, "Additional Notes/Comments");

            foreach (var note in consts_parser.AdditionalNotes)
            {
                ++row_num;

                sl.SelectWorksheet(ConstraintsReferenceParser.ReferenceSheetName);
                SLStyle source_style = sl.GetCellStyle(note);
                source_style.Alignment.Horizontal = HorizontalAlignmentValues.Left;
                source_style.Alignment.Vertical = VerticalAlignmentValues.Top;

                sl.SelectWorksheet(SHEET_NAME);
                var cell_to_write = SLConvert.ToCellReference(row_num, STRUCTURE_NAME_COL);
                sl.CopyCellFromWorksheet(ConstraintsReferenceParser.ReferenceSheetName, note, note, cell_to_write);
                var num_of_lines = sl.GetCellValueAsString(cell_to_write).Count(c => c == '\n');
                num_of_lines = (num_of_lines > 0) ? num_of_lines : 1;

                sl.SetRowHeight(row_num, CONSTS_HEIGHT * num_of_lines);
                sl.MergeWorksheetCells(row_num, STRUCTURE_NAME_COL, row_num, DECISION_COL, full_style);
                sl.SetCellStyle(cell_to_write, source_style);
            }
            for (int i = 0; i < 1; ++i)
            {
                ++row_num;
                sl.SetRowHeight(row_num, CONSTS_HEIGHT);
                sl.MergeWorksheetCells(row_num, STRUCTURE_NAME_COL, row_num, DECISION_COL, full_style);
            }

            row_num += 2;
            sl.MergeWorksheetCells(row_num, STRUCTURE_NAME_COL, row_num, STRUCTURE_NAME_COL + LABEL_WID - 1, full_style);
            sl.MergeWorksheetCells(row_num, STRUCTURE_NAME_COL + LABEL_WID, row_num, STRUCTURE_NAME_COL + LABEL_WID + VALUE_WID - 1, full_style);
            sl.SetCellValue(row_num, STRUCTURE_NAME_COL, "Plan start date");

            sl.SetFontToAllCells(FONT_NAME);
            sl.SetPrintArea(1, 1, sl.GetWorksheetStatistics().EndRowIndex, sl.GetWorksheetStatistics().EndColumnIndex);
            //sl.Save();
            sl.SaveAs(output_file_path);

        }


        private void WriteHeaderProperty(SLDocument sl, string start_cell_ref, string label, string value, int label_wid, int value_wid)
        {
            (int, int) start_pos = SpreadSheetLightExtensions.GetRowColIndexFromCellReference(start_cell_ref);
            (int, int) value_pos = (start_pos.Item1, start_pos.Item2 + label_wid);

            sl.MergeWorksheetCells(start_pos.Item1, start_pos.Item2, start_pos.Item1, value_pos.Item2 - 1, header_style_index);

            sl.MergeWorksheetCells(value_pos.Item1, value_pos.Item2, value_pos.Item1, value_pos.Item2 + value_wid - 1, header_style_value);

            sl.SetCellValue(start_cell_ref, label);
            sl.SetCellStyle(start_cell_ref, header_style_index);

            sl.SetCellValue(value_pos.Item1, value_pos.Item2, value);
            sl.SetCellStyle(value_pos.Item1, value_pos.Item2, header_style_value);


            var num_of_lines = value.Count(c => c == '\n') + 1;

            sl.SetRowHeight(start_pos.Item1, HEADER_HEIGHT * num_of_lines);

            return;
        }


        private void WriteCIRow(SLDocument sl, ref int row_num,
            string volume_name, string symbol, double volume)
        {
            sl.SetRowHeight(row_num, CI_ROW_HEIGHT);


            // the name of volume
            sl.SetCellValue(row_num, STRUCTURE_NAME_COL, volume_name);
            sl.SetCellValue(row_num, STRUCTURE_NAME_COL, volume_name);

            start_style.SetWrapText(true);
            sl.MergeWorksheetCells(row_num, STRUCTURE_NAME_COL, row_num, DVH_INDEX_COL, start_style);
            start_style.SetWrapText(false);

            // the symbol of volume
            sl.SetCellValue(row_num, RELATION_COL, symbol);
            sl.MergeWorksheetCells(row_num, RELATION_COL, row_num, ACCEPTABLE_INDEX_VALUE_COL, mid_style);

            // volume
            sl.SetCellValue(row_num, ACCEPTABLE_UNIT_COL, volume);

            mid_style.FormatCode = "0.00";
            sl.MergeWorksheetCells(row_num, ACCEPTABLE_UNIT_COL, row_num, ACTUAL_DVH_INDEX_COL, mid_style);
            mid_style.RemoveFormatCode();

            // unit
            sl.SetCellValueAndStyle(row_num, EQUAL_COL, "cc", end_style);

            ++row_num;

            return;
        }


        private void WriteConstsRowsInSummary(SLDocument sl, Int32 row, string st_name, List<DoseConstraints> consts)
        {
            sl.SetCellValue(row, STRUCTURE_NAME_COL, st_name);
            sl.MergeWorksheetCells(row, STRUCTURE_NAME_COL, row + consts.Count - 1, STRUCTURE_NAME_COL/*, full_style*/);

            Int32 count = 0;
            foreach (DoseConstraints con in consts)
            {
                // CellStyle of STRUCTURE_NAME_COL must be set here.
                // If SetCellStyle() is not called at following rows, 
                // cell style is not applied if merged row is not exist (only 1 row is exist).
                sl.SetCellStyle(row + count, STRUCTURE_NAME_COL, full_style);

                /* translate DoseConstraints */
                var enum_strings = new EnumExtensions.Enums2Strings(con);

                // decision
                string tolerance_cell = SLConvert.ToCellReference(row + count, TOLERANCE_INDEX_VALUE_COL);
                string acceptable_cell = SLConvert.ToCellReference(row + count, ACCEPTABLE_INDEX_VALUE_COL);
                string actual_cell = SLConvert.ToCellReference(row + count, ACTUAL_INDEX_VALUE_COL);
                string inequality_sign = (enum_strings.Relationship == "≦") ? "<="
                                            : (enum_strings.Relationship == "≧") ? ">=" : enum_strings.Relationship;
                string decision_formula = string.Format("=IF(ISBLANK({0}),\"\",IF({0}{1}{2},\"○\",IF({0}{1}{3},\"△\",\"×\")))", actual_cell, inequality_sign, tolerance_cell, acceptable_cell);

                /* Write to file */
                string format_code = "0.0";
                if (enum_strings.Index == "CI")
                {
                    format_code = "0.00";
                }
                else {; }

                sl.SetRowHeight(row + count, CONSTS_HEIGHT);

                if (con.Relationship == DoseConstraints.MagnitudeRelationship.Equal)
                {
                    start_style.Fill.SetPatternType(PatternValues.Solid);
                    start_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xfcd5b5));
                    mid_style.Fill.SetPatternType(PatternValues.Solid);
                    mid_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xfcd5b5));
                    end_style.Fill.SetPatternType(PatternValues.Solid);
                    end_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xfcd5b5));
                }
                else {; }

                sl.SetCellValueAndStyle(row + count, DVH_INDEX_COL, enum_strings.Index, start_style);
                sl.SetCellValueAndStyle(row + count, RELATION_COL, enum_strings.Relationship, mid_style);
                if (con.Tolerance != null)
                {
                    sl.SetCellValueAndStyle(row + count, TOLERANCE_INDEX_VALUE_COL, con.Tolerance.Value, mid_style, format_code);
                }
                else
                {
                    sl.SetCellValueAndStyle(row + count, TOLERANCE_INDEX_VALUE_COL, "", mid_style);
                }
                sl.SetCellValueAndStyle(row + count, TOLERANCE_UNIT_COL, enum_strings.Unit, end_style);
                if (con.Acceptable != null)
                {
                    sl.SetCellValueAndStyle(row + count, ACCEPTABLE_INDEX_VALUE_COL, con.Acceptable.Value, start_style, format_code);
                }
                else
                {
                    sl.SetCellValueAndStyle(row + count, ACCEPTABLE_INDEX_VALUE_COL, "", start_style);
                }
                sl.SetCellValueAndStyle(row + count, ACCEPTABLE_UNIT_COL, enum_strings.Unit, end_style);

                start_style.Fill.SetPatternType(PatternValues.Solid);
                start_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xfcd5b5));
                mid_style.Fill.SetPatternType(PatternValues.Solid);
                mid_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xfcd5b5));
                end_style.Fill.SetPatternType(PatternValues.Solid);
                end_style.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xfcd5b5));
                sl.SetCellValueAndStyle(row + count, ACTUAL_DVH_INDEX_COL, enum_strings.Index, start_style);
                sl.SetCellValueAndStyle(row + count, EQUAL_COL, "=", mid_style);

                if (con.ActualValue.HasValue == false)
                {
                    sl.SetCellValueAndStyle(row + count, ACTUAL_INDEX_VALUE_COL, "", mid_style);
                }
                else if (double.IsNaN(con.ActualValue.Value) == true)
                {
                    sl.SetCellValueAndStyle(row + count, ACTUAL_INDEX_VALUE_COL, "NaN", mid_style);
                }
                else 
                {
                    sl.SetCellValueAndStyle(row + count, ACTUAL_INDEX_VALUE_COL, con.ActualValue.Value, mid_style, format_code);
                }
                sl.SetCellValueAndStyle(row + count, ACTUAL_UNIT_COL, enum_strings.Unit, end_style);
                start_style.Fill.SetPatternType(PatternValues.None);
                mid_style.Fill.SetPatternType(PatternValues.None);
                end_style.Fill.SetPatternType(PatternValues.None);

                full_style.Font.FontSize = DECISION_FONT_SIZE;
                sl.SetCellValueAndStyle(row + count, DECISION_COL, decision_formula, full_style);
                full_style.Font.FontSize = CONTENTS_FONT_SIZE;

                ++count;

            }

            return;
        }

    }

    internal static class SpreadSheetLightExtensions
    {

        public static void SetCellValueAndStyle(this SLDocument sl, int RowIndex, int ColumnIndex, string CellValue, SLStyle CellStyle)
        {
            sl.SetCellValue(RowIndex, ColumnIndex, CellValue);
            sl.SetCellStyle(RowIndex, ColumnIndex, CellStyle);

            return;
        }

        public static void SetCellValueAndStyle(this SLDocument sl, int RowIndex, int ColumnIndex, double CellValue, SLStyle CellStyle, string format_code = "0.0")
        {
            CellStyle.FormatCode = format_code;
            sl.SetCellValue(RowIndex, ColumnIndex, CellValue);
            sl.SetCellStyle(RowIndex, ColumnIndex, CellStyle);

            return;
        }

        public static (int, int) GetRowColIndexFromCellReference(string cellReference)
        {
            string rowPart = new string(cellReference.Where(char.IsDigit).ToArray());
            int rowIndex = int.Parse(rowPart);
            int columnIndex = SLConvert.ToColumnIndex(cellReference);

            return (rowIndex, columnIndex);
        }

        public static void SetFontToAllCells(this SLDocument sl, string font_name)
        {
            for (int row = 1; row <= sl.GetWorksheetStatistics().EndRowIndex; row++)
            {
                for (int col = 1; col <= sl.GetWorksheetStatistics().EndColumnIndex; col++)
                {
                    var original_style = sl.GetCellStyle(row, col);
                    var original_font = original_style.Font;
                    var font_size = (original_font.FontSize.HasValue) ? original_font.FontSize.Value : 11;
                    original_style.SetFont(font_name, font_size);

                    sl.SetCellStyle(row, col, original_style);
                }
            }

            return;
        }


    }
}
