using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
//using System.Reactive;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using static constraints_summarizer_esapi_v15_5.DoseConstraints;

namespace constraints_summarizer_esapi_v15_5
{
    internal class DoseConstraints
    {
        // ==============================
        //       Enum & Class Declaration 
        // ==============================
        public enum DvhIndex
        {
            D_max,
            D_min,
            D_mean,
            D_percent,
            D_cc,
            DC_percent,
            DC_cc,
            V_percent,
            V_Gy,
            CV_percent,
            CV_Gy,
            CI,
        }

        public static bool IsDoseMaxMinMeanConsts(DvhIndex index)
        {
            bool res = false;
            if ((index == DvhIndex.D_max)
                || (index == DvhIndex.D_min)
                || (index == DvhIndex.D_mean))
            {
                res = true;
            }
            else
            {
                res = false;
            }

            return res;
        }

        public static bool IsDoseConsts(DvhIndex index)
        {
            bool res = false;
            if ((IsDoseMaxMinMeanConsts(index) == true)
                || (index == DvhIndex.D_percent)
                || (index == DvhIndex.D_cc)
                || (index == DvhIndex.DC_percent)
                || (index == DvhIndex.DC_cc))
            {
                res = true;
            }
            else
            {
                res = false;
            }

            return res;
        }

        public static bool IsCIConsts(DvhIndex index)
        {
            bool res = false;
            if (index == DvhIndex.CI)
            {
                res = true;
            }
            else
            {
                res = false;
            }

            return res;
        }

        public static bool IsVolumeConsts(DvhIndex index)
        {
            return (!IsDoseConsts(index)) & (!IsCIConsts(index));
        }


        public enum MagnitudeRelationship
        {
            Equal,
            //            NotEqual, 
            LessThan,
            LessThanOrEqual,
            GreaterThan,
            GreaterThanOrEqual

        }

        public enum AbsRel
        {
            Absolute,
            Relative,
            Dimensionless,
        }

        private static DvhIndex ParseDvhIndex(in string input)
        {
            switch (input)
            {
                case "Dmax": return DvhIndex.D_max;
                case "Dmin": return DvhIndex.D_min;
                case "Dmean": return DvhIndex.D_mean;
                case "Dx%": return DvhIndex.D_percent;
                case "Dxcc": return DvhIndex.D_cc;
                case "DCx%": return DvhIndex.DC_percent;
                case "DCxcc": return DvhIndex.DC_cc;
                case "Vx%": return DvhIndex.V_percent;
                case "VxGy": return DvhIndex.V_Gy;
                case "CVx%": return DvhIndex.CV_percent;
                case "CVxGy": return DvhIndex.CV_Gy;
                case "CI": return DvhIndex.CI;
                default:
                    throw new ArgumentException(string.Format("Input: {0} is not an appropriate value for DvhIndex.\n", input));
            }
        }

        private static MagnitudeRelationship ParseMagnitudeRelationship(in string input)
        {
            switch (input)
            {
                case "=": return MagnitudeRelationship.Equal;
                case "<": return MagnitudeRelationship.LessThan;
                case "≦": return MagnitudeRelationship.LessThanOrEqual;
                case ">": return MagnitudeRelationship.GreaterThan;
                case "≧": return MagnitudeRelationship.GreaterThanOrEqual;
                default:
                    throw new ArgumentException(string.Format("Input: {0} is not an appropriate value for MagnitudeRelationship.\n", input));
            }
        }

        private static AbsRel ParseAbsRel(in string input)
        {
            switch (input)
            {
                case "Gy": return AbsRel.Absolute;
                case "cc": return AbsRel.Absolute;
                case "%": return AbsRel.Relative;
                case "": return AbsRel.Dimensionless;
                default:
                    throw new ArgumentException(string.Format("Input: {0} is not an appropriate value for AbsRel.\n", input));
            }
        }

        public class ForCICalc
        {
            public double PrescriptionDoseGy { get; set; } = 0.0;
            public double TargetVolumeCc { get; set; } = 0.0;
            public double PtvIrradiatedVolumeCc { get; set; } = 0.0;
            public double AllIrradiatedVolumeCc { get; set; } = 0.0;
        }

        // ==============================
        //       Fields Declaration 
        // ==============================
        // IndexValue, Tolerance and Acecptable are omittable.
        // (Dmax/Dmin/Dmean don't require IndexValue.)
        // (In some cases, only one of Tolerance and Acceptable is specified.)
        static public string GlobalDoseMax = "3D Dose MAX";
        public string Structure { get; private set; }
        public DvhIndex Index { get; private set; }
        public double? IndexValue { get; private set; }
        public MagnitudeRelationship Relationship { get; private set; }
        public double? Tolerance { get; private set; }
        public double? Acceptable { get; private set; }
        public AbsRel Unit { get; private set; }
        public double? ActualValue { get; set; } = null;

        public ForCICalc? ForCI { get; private set; } = null;

        private DoseConstraints(string structure,
            DvhIndex index
            , double? indexValue
            , MagnitudeRelationship relationship
            , double? tolerance
            , double? acceptable
            , AbsRel unit
            , double? actual
            , ForCICalc? ci
            )
        {
            Structure = structure ?? throw new ArgumentNullException(nameof(structure));
            Index = index;
            IndexValue = indexValue;
            Relationship = relationship;
            Tolerance = tolerance;
            Acceptable = acceptable;
            Unit = unit;
            ActualValue = actual;
            ForCI = ci;
        }

        // This constructor is used to parse an input "Constraints Sheet".
        // Input List<string> is a line of the sheet.
        public DoseConstraints(in List<string> strings)
        {

            if (strings == null)
            {
                throw new ArgumentNullException("ERROR: The argument is null.\n");
            }
            else if (strings.Count < 7)
            {
                // If IndexValue and Acceptable or Toleranve are null, strings.Count == 5 
                throw new ArgumentException("ERROR: The Count() of the argument should be larger than or equal to 7.\n");
            }
            else
            {
                try
                {
                    double res = 0.0;

                    Structure = strings.ElementAt(0);
                    Index = ParseDvhIndex(strings.ElementAt(1));
                    IndexValue = double.TryParse(strings.ElementAt(2), out res) ? (double?)res : null;
                    Relationship = ParseMagnitudeRelationship(strings.ElementAt(3));
                    Tolerance = double.TryParse(strings.ElementAt(4), out res) ? (double?)res : null;
                    Acceptable = double.TryParse(strings.ElementAt(5), out res) ? (double?)res : null;
                    Unit = ParseAbsRel(strings.ElementAt(6));

                    // fields below are fields not parsed from the input "Constraints Sheet"
                    ActualValue = null;
                    ForCI = (Index == DvhIndex.CI) ? new ForCICalc() : null;
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("ERROR: Exception occurred at DoseConstraints(in List<string> strings.\nStackTrace: {0}\nInnerException: {1}\nHelpLink: {2}\nMessage: {3}\n", ex.StackTrace, ex.InnerException, ex.HelpLink, ex.Message));
                }
            }
        }

        public string DetailToString()
        {
            string buf = "";

            buf += string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}\n"
                , Structure
                , Index.ToString()
                , IndexValue.HasValue ? IndexValue.Value : "null"
                , Relationship.ToString()
                , Tolerance.HasValue ? Tolerance.Value : "null"
                , Acceptable.HasValue ? Acceptable.Value : "null"
                , Unit.ToString());

            return buf;
        }

        public static Comparison<DoseConstraints> Compare = (lhs, rhs) =>
        {

            int ret_lhs = 0;
            int ret_rhs = 0;
            int ret = 0;

            /* Step 1: Structure name */
            if (lhs.Structure != rhs.Structure)
            {
                ret = lhs.Structure.CompareTo(rhs.Structure);
            }
            /* Step 2: CI (Conformity Index), Dose (Dmax/min/mean or D_others) or Volume */
            else
            {
                Func<DoseConstraints, int> DvhIndexGroup = (constraints) =>
                {
                    int ret_group = 0;

                    if (IsDoseMaxMinMeanConsts(constraints.Index))
                    {
                        ret_group = 10;
                    }
                    else if (IsVolumeConsts(constraints.Index))
                    {
                        ret_group = 30;
                    }
                    else if (IsCIConsts(constraints.Index))
                    {
                        ret_group = 0;
                    }
                    else
                    {
                        ret_group = 20;
                    }

                    return ret_group;
                };

                ret_lhs = DvhIndexGroup(lhs);
                ret_rhs = DvhIndexGroup(rhs);

                // If the DvhIndex of the lhs is different from that of the rhs,
                // the ActualValue comparison can be ommited.
                if (ret_lhs != ret_rhs)
                {; }
                else
                {
                    /* Step 3: Actual Value */
                    double actual_lhs = lhs.ActualValue.GetValueOrDefault();
                    double actual_rhs = rhs.ActualValue.GetValueOrDefault();
                    int ret_actual = 0;

                    if ((ret_lhs == 10) || (ret_lhs == 20))  // if it's Dose Index
                    {
                        // rescale the return value of ComparaTo() to -1, 0, 1
                        ret_actual = (actual_lhs.CompareTo(actual_rhs) == 0) ? 0
                            : (actual_rhs.CompareTo(actual_lhs) > 0) ? 1
                            : -1;
                    }
                    else if (ret_lhs == 0)// if it's CI (Conformity Index)
                    {
                        ret_actual = 0;
                    }
                    else  // if it's Volume Index
                    {
                        // INVERT the result for descending ordering.
                        ret_actual = (actual_lhs.CompareTo(actual_rhs) == 0) ? 0
                            : (actual_rhs.CompareTo(actual_lhs) > 0) ? -1
                            : 1;

                    }

                    ret_lhs += ret_actual;
                }

                ret = ret_lhs.CompareTo(ret_rhs);
            }

            return ret;
        };
    }


    internal class DoseConstraintsDict
    {
        // ============================
        //      Fields declaration
        // ============================
        public Dictionary<string, List<DoseConstraints>> Dict { get; private set; } = new Dictionary<string, List<DoseConstraints>>();

        public DoseConstraintsDict(in List<DoseConstraints> consts)
        {
            foreach (DoseConstraints d in consts)
            {
                this.Add(d);
            }

            foreach (var consts_pair in this.Dict)
            {
                consts_pair.Value.Sort(DoseConstraints.Compare);
            }
        }


        public void Add(DoseConstraints value)
        {
            string key = value.Structure;

            // Update dict
            if (Dict.ContainsKey(key) == false)
            {
                Dict[key] = new List<DoseConstraints>();
            }
            else {; }

            Dict[key].Add(value);

            return;
        }

    }

    internal static class EnumExtensions
    {
        public static string RelationToString(this MagnitudeRelationship rel)
        {
            string ret = "";
            switch (rel)
            {
                case MagnitudeRelationship.Equal: ret = "="; break;
                case MagnitudeRelationship.LessThan: ret = "<"; break;
                case MagnitudeRelationship.LessThanOrEqual: ret = "≦"; break;
                case MagnitudeRelationship.GreaterThan: ret = ">"; break;
                case MagnitudeRelationship.GreaterThanOrEqual: ret = "≧"; break;
                default: ret = ""; break;
            }

            return ret;
        }

        public class Enums2Strings
        {
            public string Index { get; private set; }
            public string Relationship { get; private set; }
            public string Unit { get; private set; }
            public string Decision { get; private set; }

            public Enums2Strings(DoseConstraints con)
            {
                double? idx_val = con.IndexValue;
                string index = "";
                if ((idx_val == null) &&
                        ((IsDoseMaxMinMeanConsts(con.Index) == false) && (IsCIConsts(con.Index) == false)))
                {; }
                else
                {
                    switch (con.Index)
                    {
                        case DvhIndex.D_max: index = "Dmax"; break;
                        case DvhIndex.D_min: index = "Dmin"; break;
                        case DvhIndex.D_mean: index = "Dmean"; break;
                        case DvhIndex.D_percent: index = string.Format("D{0}%", idx_val.Value); break;
                        case DvhIndex.D_cc: index = string.Format("D{0}cc", idx_val.Value); break;
                        case DvhIndex.DC_percent: index = string.Format("DC{0}%", idx_val.Value); break;
                        case DvhIndex.DC_cc: index = string.Format("DC{0}cc", idx_val.Value); break;
                        case DvhIndex.V_percent: index = string.Format("V{0}%", idx_val.Value); break;
                        case DvhIndex.V_Gy: index = string.Format("V{0}Gy", idx_val.Value); break;
                        case DvhIndex.CV_percent: index = string.Format("CV{0}%", idx_val.Value); break;
                        case DvhIndex.CV_Gy: index = string.Format("CV{0}Gy", idx_val.Value); break;
                        case DvhIndex.CI: index = "CI"; break;
                        default: break;
                    }
                }

                string inequality_sign = con.Relationship.RelationToString();

                string unit = "";
                if (con.Unit == AbsRel.Relative)
                {
                    unit = "%";
                }
                else if (IsDoseConsts(con.Index))
                {
                    unit = "Gy";
                }
                else if (IsVolumeConsts(con.Index))
                {
                    unit = "cc";
                }
                else
                {
                    unit = "";
                }

                // decision
                string decision = (con.ActualValue.HasValue == false) ? "-"
                             : IsRelationshipSatisfied(con.Relationship, con.ActualValue.Value, con.Tolerance) ? "〇"
                             : IsRelationshipSatisfied(con.Relationship, con.ActualValue.Value, con.Acceptable) ? "△" : "×";


                /* return */
                Index = index;
                Relationship = inequality_sign;
                Unit = unit;
                Decision = decision;
            }
        }

        static public bool IsRelationshipSatisfied(DoseConstraints.MagnitudeRelationship relation, double actual, double? ref_nullable)
        {
            bool ret = false;

            if (ref_nullable.HasValue == false)
            {
                ret = false;
            }
            else
            {
                double reference = ref_nullable.Value;

                switch (relation)
                {
                    case DoseConstraints.MagnitudeRelationship.Equal:
                        {
                            ret = (AreAlmostEqual(actual, reference, 4)) ? true : false;
                            break;
                        }
                    case DoseConstraints.MagnitudeRelationship.LessThan:
                        {
                            ret = (actual < reference) ? true : false;
                            break;
                        }
                    case DoseConstraints.MagnitudeRelationship.LessThanOrEqual:
                        {
                            ret = (AreAlmostEqual(actual, reference, 4)) ? true
                                : (actual < reference) ? true : false;
                            break;
                        }
                    case DoseConstraints.MagnitudeRelationship.GreaterThan:
                        {
                            ret = (actual > reference) ? true : false;
                            break;
                        }
                    case DoseConstraints.MagnitudeRelationship.GreaterThanOrEqual:
                        {
                            ret = (AreAlmostEqual(actual, reference, 4)) ? true
                                : (actual > reference) ? true : false;
                            break;
                        }
                }
            }

            return ret;
        }

        public static bool AreAlmostEqual(double a, double b, int precision)
        {
            // 丸めのために10のprecision乗をかけ、その後で割る
            double multiplier = Math.Pow(10, precision);

            // aとbを小数点以下precision位で丸める
            double roundedA = Math.Round(a * multiplier) / multiplier;
            double roundedB = Math.Round(b * multiplier) / multiplier;

            // 丸めた結果を比較
            return roundedA == roundedB;
        }
    }

}
