using constraints_summarizer_esapi_v15_5;
using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static constraints_summarizer_esapi_v15_5.DoseConstraints;

using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace constraints_summarizer_esapi_v15_5
{
    internal class ViewModel
    {
        public ConstraintsReferenceParser ConstsParser { get; private set; } = new ConstraintsReferenceParser();
        public ConstraintsFormGen ConstsFormGen { get; private set; }

        public string PatientId { get; private set; }
        public string PatientName { get; private set; }
        public string PlanName { get; private set; }
        public string Energy { get; private set; }
        public string Machine { get; private set; }
        public string CalcAlgorithm { get; private set; }
        public string CalcGridSize { get; private set; }
        public string PlanNormalizationMethod { get; private set; }

        public void ReadConstraintsReference(in string ref_file_path)
        {
            var error = ConstsParser.Parse(ref_file_path);
            if (error != null) {
                MessageBox.Show(error);
            }
            else {; }

            ConstsParser.Constraints.Sort(DoseConstraints.Compare);
            //ConstsParser.Print();

            // for preview
            if (ConstsParser.Constraints == null)
            {
                throw new NullReferenceException("Constraints reference might not be parsed.");
            }
            else
            {; }

            return;
        }

        private (DoseValuePresentation, VolumePresentation) ConstraintsToDosePresentation(DoseConstraints con)
        {
            var d_abs = DoseValuePresentation.Absolute;
            var d_rel = DoseValuePresentation.Relative;
            var v_abs = VolumePresentation.AbsoluteCm3;
            var v_rel = VolumePresentation.Relative;

            DoseValuePresentation dp = d_abs;
            VolumePresentation vp = v_abs;

            
            switch (con.Index)
            {
                case DvhIndex.D_max: dp = (con.Unit == AbsRel.Absolute) ? d_abs : d_rel; break;
                case DvhIndex.D_min: dp = (con.Unit == AbsRel.Absolute) ? d_abs : d_rel; break;
                case DvhIndex.D_mean: dp = (con.Unit == AbsRel.Absolute) ? d_abs : d_rel; break;

                case DvhIndex.D_percent: { 
                        dp = (con.Unit == AbsRel.Absolute) ? d_abs : d_rel; 
                        vp = v_rel; 
                        break; 
                } 
                case DvhIndex.D_cc: { 
                        dp = (con.Unit == AbsRel.Absolute) ? d_abs : d_rel; 
                        vp = v_abs; 
                        break; 
                } 
                case DvhIndex.DC_percent: { 
                        dp = (con.Unit == AbsRel.Absolute) ? d_abs : d_rel; 
                        vp = v_rel; 
                        break; 
                } 
                case DvhIndex.DC_cc: { 
                        dp = (con.Unit == AbsRel.Absolute) ? d_abs : d_rel; 
                        vp = v_abs; 
                        break; 
                } 

                case DvhIndex.V_percent: { 
                        dp = d_rel; 
                        vp = (con.Unit == AbsRel.Absolute) ? v_abs : v_rel; 
                        break; 
                } 
                case DvhIndex.V_Gy: { 
                        dp = d_abs; 
                        vp = (con.Unit == AbsRel.Absolute) ? v_abs : v_rel; 
                        break; 
                } 
                case DvhIndex.CV_percent: { 
                        dp = d_rel; 
                        vp = (con.Unit == AbsRel.Absolute) ? v_abs : v_rel; 
                        break; 
                } 
                case DvhIndex.CV_Gy: { 
                        dp = d_abs; 
                        vp = (con.Unit == AbsRel.Absolute) ? v_abs : v_rel; 
                        break; 
                } 

                case DvhIndex.CI: break;

                default: break;
            }

            return (dp, vp);
        }

        public void AcquireActualDoses(ScriptContext context, in List<DoseConstraints> consts, ExternalPlanSetup plan_setup, PlanSum plan_sum)
        {
            PlanningItem plan = (plan_sum != null) ? plan_sum : plan_setup;

            var structures = plan.StructureSet.Structures;
            double bin_width = 0.01;
            List<string> energy_list = new List<string>();
            
            /* Patient information */
            PatientId = context.Patient.Id;
            PatientName = $"{context.Patient.LastName} {context.Patient.FirstName}";
            PlanName = plan.Id;
            Energy = "";
            if (plan_sum != null)
            {
                var plan_setups = plan_sum.PlanSetups;
                Machine = "";
                CalcAlgorithm = "";
                CalcGridSize = "";
                PlanNormalizationMethod = "";

                int i = 0;
                foreach (var ps in plan_sum.PlanSetups.OrderBy(x => x.CreationDateTime))
                {
                    energy_list.Add("");
                    foreach (Beam beam in ps.Beams)
                    {
                        if (energy_list[i].Contains(beam.EnergyModeDisplayName) == false)
                        {
                            energy_list[i] += (energy_list[i] == "") ? "" : ", ";
                            energy_list[i] += beam.EnergyModeDisplayName;
                        }
                        else {; }
                    }
                    Energy = string.Join("\n", energy_list);

                    Machine += ((i > 0) ? "\n" : "") + ps.Beams.ElementAt(0).TreatmentUnit.Id;
                    PlanName += "\n - " + ps.Id;
                    CalcAlgorithm += ((i > 0) ? "\n" : "") + ps.PhotonCalculationModel;
                    CalcGridSize += ((i > 0) ? "\n" : "") + ps.PhotonCalculationOptions["CalculationGridSizeInCM"];
                    PlanNormalizationMethod += ((i > 0) ? "\n" : "") + ps.PlanNormalizationMethod;

                    ++i;
                }
            }
            else {
                foreach (Beam beam in plan_setup.Beams)
                {
                    if (Energy.Contains(beam.EnergyModeDisplayName) == false)
                    {
                        Energy += (Energy != "") ? ", " : "";
                        Energy += beam.EnergyModeDisplayName;
                    }
                    else {; }
                }

                Machine = plan_setup.Beams.ElementAt(0).TreatmentUnit.Id;
                CalcAlgorithm = plan_setup.PhotonCalculationModel;
                CalcGridSize = plan_setup.PhotonCalculationOptions["CalculationGridSizeInCM"];
                PlanNormalizationMethod = plan_setup.PlanNormalizationMethod;
            }


            foreach (DoseConstraints d in consts)
            {
                /* Replace "GlobalDoseMax" to the "maximum dose to Body" */
                Structure structure = null;
                if (d.Structure == DoseConstraints.GlobalDoseMax)
                {
                    structure = structures.Where(x => x.DicomType == "EXTERNAL").ElementAtOrDefault(0);
                }
                else
                {
                    structure = structures.Where(x => x.Id == d.Structure).ElementAtOrDefault(0);
                }

                if ((structure != null))
                {
                    try
                    {
                        if (structure.IsEmpty)
                        {
                            throw new Exception(String.Format("{0}が空です。\n", structure.Id));
                        }
                        var (dp, vp) = ConstraintsToDosePresentation(d);

                        if (DoseConstraints.IsDoseMaxMinMeanConsts(d.Index))
                        {
                            var cumulative_data = plan.GetDVHCumulativeData(structure, dp, vp, bin_width);

                            if (d.Index == DvhIndex.D_min)
                            {
                                d.ActualValue = cumulative_data.MinDose.Dose;
                            }
                            else if (d.Index == DvhIndex.D_max)
                            {
                                d.ActualValue = cumulative_data.MaxDose.Dose;
                            }
                            else
                            {
                                d.ActualValue = cumulative_data.MeanDose.Dose;
                            }

                            if (dp == DoseValuePresentation.Absolute)
                            {
                                // convert cGy (system units) to Gy
                                d.ActualValue /= 100.0;
                            }
                            else {; }
                        }
                        else if (DoseConstraints.IsVolumeConsts(d.Index))
                        {
                            var dose_unit = (dp == DoseValuePresentation.Absolute) ? DoseValue.DoseUnit.cGy : DoseValue.DoseUnit.Percent;
                            var index_value = (dp == DoseValuePresentation.Absolute) ? 
                                                d.IndexValue.Value * 100 /* convert Gy to cGy */ : d.IndexValue.Value;

                            d.ActualValue = plan.GetVolumeAtDose(structure, new DoseValue(index_value, dose_unit), vp);

                            /* CV_x[cc/%] = (Volume of whole organ) - (GetVolumeAtDose(x)) */
                            if ((d.Index == DvhIndex.CV_Gy) || (d.Index == DvhIndex.CV_percent))
                            {
                                var whole_volume = (vp == VolumePresentation.Relative) ? 100.0 : structure.Volume;
                                d.ActualValue = whole_volume - d.ActualValue;
                            }
                            else {; }
                        }
                        else if (DoseConstraints.IsCIConsts(d.Index))
                        {
                            Structure body = structures.Where(x => x.DicomType == "EXTERNAL").ElementAtOrDefault(0);

                            if (body == null)
                            {
                                throw new Exception("ERROR: Structure set doesn't have BODY.");
                            }
                            else
                            {
                                d.ForCI.PrescriptionDoseGy = ConstsParser.TotalDose * (ConstsParser.PrescriptionPercentage * 0.01);
                                d.ForCI.TargetVolumeCc = structure.Volume;
                                d.ForCI.PtvIrradiatedVolumeCc = plan.GetVolumeAtDose(structure, new DoseValue(d.ForCI.PrescriptionDoseGy * 100, "cGy"), vp);
                                d.ForCI.AllIrradiatedVolumeCc = plan.GetVolumeAtDose(body, new DoseValue(d.ForCI.PrescriptionDoseGy * 100, "cGy"), vp);

                                d.ActualValue = Math.Pow(d.ForCI.PtvIrradiatedVolumeCc, 2)
                                                    / (d.ForCI.TargetVolumeCc * d.ForCI.AllIrradiatedVolumeCc);
                            }
                        }
                        else
                        {
                            /* DC_x[Gy/%] = GetDoseAtVolume((Volume of whole organ) - x) */
                            if ((d.Index == DvhIndex.DC_cc) || (d.Index == DvhIndex.DC_percent))
                            {
                                var whole_volume = (vp == VolumePresentation.Relative) ? 100.0 : structure.Volume;
                                double sub_volume = whole_volume - d.IndexValue.Value;

                                if (Math.Sign(sub_volume) == -1)        // if sub_volume is negative value => can not be defined!
                                {
                                    throw new Exception(String.Format("{0}のDC制約の体積が不正です (sub_volumeの体積が負になってしまいます)。\n", structure.Id));
                                }
                                else
                                {
                                    d.ActualValue = plan.GetDoseAtVolume(structure, sub_volume, vp, dp).Dose;
                                }
                            }
                            else {
                                d.ActualValue = plan.GetDoseAtVolume(structure, d.IndexValue.Value, vp, dp).Dose;
                            }

                            if (dp == DoseValuePresentation.Absolute)
                            {
                                // convert cGy (system units) to Gy
                                d.ActualValue /= 100.0;
                            }
                            else {; }
                        }

                    }
                    catch(Exception e)
                    {
                        var (dp, vp) = ConstraintsToDosePresentation(d);
                        string buf = $"ERROR: Actual Dose計算時にエラーが発生しました。\nStructure name {structure.Id}, DVHIndex: {d.Index.ToString()}, IndexValue: {d.IndexValue}, DosePresentation: {dp}, VolumePresentation: {vp}";
                        MessageBox.Show($"{buf}\nException.Message: {e.Message}");
                    }

                }
                else {; }

            }

            return;
        }

        public List<ConstraintsFormPreview> GetListOfTargetConstraintsFormPreview()
        {
            var form_preview = new ConstraintsFormPreviewList(ConstsParser.ConstraintsDictTarget);
            return form_preview.preview_list;
        }

        public List<ConstraintsFormPreview> GetListOfOARConstraintsFormPreview()
        {
            var form_preview = new ConstraintsFormPreviewList(ConstsParser.ConstraintsDictOAR);
            return form_preview.preview_list;
        }


        public void WriteConstraintsSheet(string write_file_path)
        {
            if (ConstsParser.Constraints == null)
            {
                throw new NullReferenceException("Constraints reference might not be parsed.");
            }
            else
            {
                try
                {
                    var header_info = new HeaderInfo { PatientId = PatientId, Energy = Energy, Machine = Machine, PatientName = PatientName, 
                                    PlanName = PlanName, CalcAlgorithm = CalcAlgorithm, CalcGridSize = CalcGridSize,
                                    PlanNormalizationMethod = PlanNormalizationMethod};
                    ConstsFormGen = new ConstraintsFormGen(ConstsParser, header_info, write_file_path);
                    ConstsFormGen.WriteConstraintsSummarySheet();
                }
                catch
                {
                    throw new Exception("ERROR: Failed to Execute WriteConstraintsSheet(string write_file_path)");
                }



            }
        }


        static public (ExternalPlanSetup, PlanSum, string) GetPlanWithCourseIdDialog(ScriptContext context, bool sum_plan_only = false)
        {
            //string ret = null;
            (ExternalPlanSetup, PlanSum, string) ret = (null, null, null);
            var options = new List<string>();
            string selected_id = null;

            if (sum_plan_only == false)
            {
                foreach (var sum in context.PlansInScope)
                {
                    options.Add(sum.Id);
                }
            }
            else {; }

            foreach (var sum in context.PlanSumsInScope)
            {
                options.Add(sum.Id);
            }

            string message = "解析対象のPlanを選択してください。";
            var selection_dialog = new SelectPlanSumDialog(message, options);
            if (selection_dialog.ShowDialog() == true)
            {
//                    MessageBox.Show(string.Format("True: {0}", selection_dialog.SelectedItem));
                selected_id = selection_dialog.SelectedItem;
            }
            else {
//                MessageBox.Show(string.Format("False: {0}", selection_dialog.SelectedItem));
                throw new Exception("Plan選択時にエラーが発生しました。");
            }


            if (selected_id != null)
            {
                if (context.PlanSumsInScope.Any(x => x.Id == selected_id) == true)
                {
                    var selected_plan = context.PlanSumsInScope.Where(x => x.Id == selected_id).FirstOrDefault();

                    ret.Item1 = null;
                    ret.Item2 = selected_plan;
                    ret.Item3 = selected_plan.Course.Id;
                }
                else
                {
                    var selected_plan = context.ExternalPlansInScope.Where(x => x.Id == selected_id).FirstOrDefault();

                    ret.Item1 = selected_plan;
                    ret.Item3 = null;
                    ret.Item3 = selected_plan.Course.Id;
                }
            }
            else {; }

            return ret;
        }
    }
}
