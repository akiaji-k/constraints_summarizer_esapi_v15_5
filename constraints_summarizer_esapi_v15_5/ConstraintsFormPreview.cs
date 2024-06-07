using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static constraints_summarizer_esapi_v15_5.DoseConstraints;

namespace constraints_summarizer_esapi_v15_5
{
    internal class ConstraintsFormPreview //: INotifyPropertyChanged
    {
        public string Structure { get; private set; }
        public string Index { get; private set; }
        public string Relationship { get; private set; }
        public double? Tolerance { get; private set; }
        public double? Acceptable { get; private set; }
        public string Unit { get; private set; }
        public double? ActualValue { get; set; }
        public string Decision { get; private set; }

        public ConstraintsFormPreview(DoseConstraints con)
        {
            var enum_strings = new EnumExtensions.Enums2Strings(con);

            Structure = con.Structure;
            Index = enum_strings.Index;
            Relationship = enum_strings.Relationship;
            Tolerance = con.Tolerance;
            Acceptable = con.Acceptable;
            Unit = enum_strings.Unit;
            ActualValue = con.ActualValue;
            Decision = enum_strings.Decision;
        }
       
    }
    internal class ConstraintsFormPreviewList
    {
        public List<ConstraintsFormPreview> preview_list { get; private set; } = new List<ConstraintsFormPreview>();

        public ConstraintsFormPreviewList(List<DoseConstraints> cons)
        {
            foreach (var con in cons)
            {
                preview_list.Add(new ConstraintsFormPreview(con));
            }

            return;
        }

        public ConstraintsFormPreviewList(DoseConstraintsDict dict)
        {
            foreach (var structure in dict.Dict)
            {
                foreach (var con in structure.Value)
                {
                    preview_list.Add(new ConstraintsFormPreview(con));
                }

            }

            return;
        }
    }
}
