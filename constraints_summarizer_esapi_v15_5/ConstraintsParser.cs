using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;
//using System.Reactive;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows;

namespace constraints_summarizer_esapi_v15_5
{
    internal class ConstraintsReferenceParser
    {
        enum ParseType
        {
            Info,
            Consts,
            Notes,
            None,
        }

        public static string ReferenceSheetName { get; } = "Reference";
        public static string ListSheetName { get; } = "List";
#nullable enable
        public string ConstraintsRefFile { get; private set; } = "";
        public string Title { get; private set; } = "";
        public string Oncologist { get; private set; } = "";
        public string Planner { get; private set; } = "";
        public double TotalDose { get; private set; }
        public double Fraction { get; private set; }
        public double PrescriptionPercentage { get; private set; }
        public List<DoseConstraints> Constraints { get; private set; } = new List<DoseConstraints>();
        public DoseConstraintsDict? ConstraintsDictTarget { get; private set; } = null;
        public DoseConstraintsDict? ConstraintsDictOAR { get; private set; } = null;
        public List<string> AdditionalNotes { get; private set; } = new List<string>();

        public string? ErrorBuf { get; private set; } = null;

        public string? Parse(in string constraints_ref_file_path)
        {
            string full_path = Path.GetFullPath(constraints_ref_file_path);

            if (!File.Exists(full_path))
            {
                throw new FileNotFoundException("ERROR: ParseConstraintsReference()の入力ファイルが存在しません。");
            }
            else
            {
                ConstraintsRefFile = full_path;

                /* FileShare is not needed for shared(legacy) Excel file */
                SLDocument sl = new SLDocument(ConstraintsRefFile);
                //System.IO.FileStream fs = new System.IO.FileStream(
                //    constraints_ref_file_path,
                //    System.IO.FileMode.Open,
                //    System.IO.FileAccess.ReadWrite,
                //    System.IO.FileShare.ReadWrite);
                //SLDocument sl = new SLDocument(fs);


                if (sl.SelectWorksheet(ReferenceSheetName) == false)
                {
                    throw new Exception("ERROR: 制約参照ファイルに'Reference'という名前のシートが存在しません。");
                }
                else {; }

                for (int row = 1; row <= sl.GetWorksheetStatistics().EndRowIndex; row++)
                {
                    var values_of_line = ReadLineAndRemoveComment(sl, row);

                    try
                    {
                        var ptype = GetPlanningInfo(values_of_line);
                        if (ptype == ParseType.Notes)
                        {
                            AdditionalNotes.Add(SLConvert.ToCellReference(row, 3));
                        }
                        else if ((values_of_line.Count > 0) 
                            && (values_of_line.All(x => x == "") == false)      // if read row is empty
                            && (GetPlanningInfo(values_of_line) == ParseType.Consts))
                        {
                            Constraints.Add(new DoseConstraints(values_of_line));
                        }
                        else {; }
                    }
                    catch (Exception e)
                    {
                        ErrorBuf += string.Format("When reading Row'{0}', ERROR: {1}\n", row, e.Message);
                    }

                }

                sl.CloseWithoutSaving();


                // Structure type classification
                Func<string, string, bool> ContainsAndNotEmpty = (x, some) =>
                {
                    return (string.IsNullOrEmpty(x) == false) && (x.Contains(some) == true);
                };
                Func<string, string, bool> CANEAndNotMinus = (x, some) =>
                {
                    var minus = "-" + some;
                    return (ContainsAndNotEmpty(x, some) == true) && (x.Contains(minus) == false);
                };

                Func<string, bool> is_target = x =>
                {
                    if ( CANEAndNotMinus(x, "PTV")
                         || CANEAndNotMinus(x, "CTV")
                         || CANEAndNotMinus(x, "ITV")
                         || CANEAndNotMinus(x, "GTV")
                         || ContainsAndNotEmpty(x, DoseConstraints.GlobalDoseMax)
                         ) return true;
                    else return false;
                };

                var consts_target = this.Constraints.Where(x => is_target(x.Structure) == true).ToList();
                var consts_oar = this.Constraints.Where(x => is_target(x.Structure) == false).ToList();
                ConstraintsDictTarget = new DoseConstraintsDict(consts_target);
                ConstraintsDictOAR = new DoseConstraintsDict(consts_oar);
            }

            return ErrorBuf;
        }

        private List<string> ReadLineAndRemoveComment(in SLDocument sl, Int32 line_num)
        {
            var values = new List<string>();

            for (int col = 1; col <= sl.GetWorksheetStatistics().EndColumnIndex; col++)
            {
                string cell_address = SLConvert.ToColumnName(col) + line_num.ToString();
                string cell_value = sl.GetCellValueAsString(cell_address);

                if (cell_value.Contains("#"))
                {     // skip comment
                    break;
                }
                else if ((col == 1) && (string.IsNullOrWhiteSpace(cell_value) == true))
                {        // remove first empty (# could be put here)
                    continue;
                }
                else
                {
                    values.Add(cell_value);
                }

                //                Console.WriteLine($"Cell {cell_address}: {cell_value}");
            }

            return values;
        }

        private ParseType GetPlanningInfo(in List<string> strings)
        {
            ParseType res = ParseType.Info;

            if (strings == null)
            {
                res = ParseType.None;
                throw new ArgumentNullException("ERROR: Input value can not be null.");
            }
            else if (strings.Count >= 2)
            {
                switch (strings.ElementAt(0))
                {
                    case "タイトル": Title = strings.ElementAt(1); break;
                    case "担当医": Oncologist = strings.ElementAt(1); break;
                    case "プラン作成者": Planner = strings.ElementAt(1); break;
                    case "総線量": TotalDose = Double.Parse(strings.ElementAt(1)); break;
                    case "分割回数": Fraction = Double.Parse(strings.ElementAt(1)); break;
                    case "総線量のx%を処方": PrescriptionPercentage = Double.Parse(strings.ElementAt(1)); break;
                    case "ノート": res = ParseType.Notes; break;
                    default: res = ParseType.Consts; break;
                }
            }
            else
            {
                res = ParseType.None;
            }

            return res;
        }

        public void Print()
        {
            string buf = "";

            buf += string.Format("Title: {0}\n", Title);
            buf += string.Format("Oncologist: {0}\n", Oncologist);
            buf += string.Format("Planner: {0}\n", Planner);
            buf += string.Format("Total dose: {0}\n", TotalDose);
            buf += string.Format("Fraction: {0}\n", Fraction);
            buf += string.Format("Prescription percentage: {0}\n", PrescriptionPercentage);

            foreach (var con in Constraints)
            {
                buf += con.DetailToString();
            }

            buf += string.Format("Errors: {0}\n", ErrorBuf);

            System.Windows.MessageBox.Show(buf);

            return;
        }

    }
}
