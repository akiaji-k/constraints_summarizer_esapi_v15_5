using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// for SpreadsheetLight
using SpreadsheetLight;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using Microsoft.Win32;

using constraints_summarizer_esapi_v15_5;

namespace VMS.TPS
{
    /// <summary>
    /// UserControl1.xaml の相互作用ロジック
    /// </summary>
    public partial class Script : UserControl
    {
        ViewModel InstViewModel = null;
        CourseReferenceSheetDB db = null;
        private const string SETTING_FILE_PATH = "$YOUR_SETTING_FILE_PATH";
        private string db_file_path = "";
        private string constraints_reference_directory = "";
        private string output_directory = "";
        private (ExternalPlanSetup, PlanSum, string) plan_with_id = (null, null, null);
        private string id = "";
        private string course = "";
        private ScriptContext context = null;

        public Script()
        {
            // to avoid System.IO.FileLoadException: "Could not load file or assembly "System.Runtime.CompilerServices.Unsafe, Version=4.0.4.1..."
            System.AppDomain.CurrentDomain.AssemblyResolve += 
                new System.ResolveEventHandler((object sender, System.ResolveEventArgs args) =>
                {
                    var name = new System.Reflection.AssemblyName(args.Name);
                    if (name.Name == "System.Runtime.CompilerServices.Unsafe")
                    {
                        return typeof(System.Runtime.CompilerServices.Unsafe).Assembly;
                    }
                    return null;
                });


            InitializeComponent();
        }

        public void Execute(ScriptContext in_context, System.Windows.Window window)
        {
            context = in_context;

            window.Content = this;
            window.SizeChanged += (sender, args) =>
            {
                this.Height = window.ActualHeight * 0.92;
                this.Width = window.ActualWidth * 0.95;
            };

            var path_settings = new PathSettings(SETTING_FILE_PATH);
            db_file_path = path_settings.DbFilePath;
            constraints_reference_directory = path_settings.ConstraintsReferenceDirectory;
            output_directory = path_settings.OutputDirectory;
            //            MessageBox.Show(String.Format("DB file: {0}, Output dir: {1}\nis read from settings file.", db_file_path, output_directory));

            db = new CourseReferenceSheetDB(db_file_path);

            id = context.Patient.Id;
            var plansetup = context.ExternalPlanSetup;
            course = "";
            if (plansetup == null)
            {
                plan_with_id = ViewModel.GetPlanWithCourseIdDialog(context, true);
                course = plan_with_id.Item3;
            }
            else
            {
                plan_with_id.Item1 = plansetup;
                course = plansetup.Course.Id;
            }

            var link = db.GetOrAddLink(id, course, constraints_reference_directory);

            if (link != null)
            {
                ShowConstraintsPreview(context, link.FilePath);
            }
            else {; }
        }

        private void ButtonClicked(object sender, RoutedEventArgs e)
        {
            // constraints processing
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Title = "Save Output File As...";
            saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls|All Files (*.*)|*.*";
            saveFileDialog.FileName = "ID_Name_Part_Dose.xlsx";
            saveFileDialog.InitialDirectory = output_directory;
            var name = InstViewModel.PatientName.Replace(" ", "_");
            saveFileDialog.FileName = $"{InstViewModel.PatientId}_{name}_部位_{InstViewModel.ConstsParser.TotalDose}.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                string file_path = saveFileDialog.FileName;

                try
                {
                    File.Copy(InstViewModel.ConstsParser.ConstraintsRefFile, file_path, true);
                    InstViewModel.WriteConstraintsSheet(file_path);
                    MessageBox.Show(string.Format("制約ファイルが生成されました。\nPath: {0}", file_path));
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"SaveFileDialog Error (エラー): {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            return;
        }

        private void ChangeSheetButtonClicked(object sender, RoutedEventArgs e)
        {
            var link = db.ModifyLink(id, course, constraints_reference_directory);

            if (link != null)
            {
                ShowConstraintsPreview(context, link.FilePath);
            }
            else {; }

            return;
        }

        private void ChangePlanButtonClicked(object sender, RoutedEventArgs e)
        {
            plan_with_id = ViewModel.GetPlanWithCourseIdDialog(context, false);
            course = plan_with_id.Item3;

            var link = db.GetOrAddLink(id, course, constraints_reference_directory);

            if (link != null)
            {
                ShowConstraintsPreview(context, link.FilePath);
            }
            else {; }

            return;
        }

        private void ShowConstraintsPreview(ScriptContext context, in string ref_file_path)
        {
            InstViewModel = new ViewModel();        // reset ViewModel

            try
            {
                InstViewModel.ReadConstraintsReference(ref_file_path);
            }
            catch (Exception _ex) 
            {
                var link = db.ModifyLink(id, course, constraints_reference_directory);
                InstViewModel.ReadConstraintsReference(link.FilePath);
            }
            InstViewModel.AcquireActualDoses(context, InstViewModel.ConstsParser.Constraints, plan_with_id.Item1, plan_with_id.Item2);

            PreviewGridTarget.ItemsSource = InstViewModel.GetListOfTargetConstraintsFormPreview();
            PreviewGridOAR.ItemsSource = InstViewModel.GetListOfOARConstraintsFormPreview();

            this.IdLabel.Content = InstViewModel.PatientId;
            this.NameLabel.Content = InstViewModel.PatientName;
            this.PlanNameLabel.Content = InstViewModel.PlanName;
            this.ReferenceSheetLabel.Text = InstViewModel.ConstsParser.ConstraintsRefFile;

            return;
        }

    }
}
