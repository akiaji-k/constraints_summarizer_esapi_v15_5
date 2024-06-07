using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json.Serialization;
using System.Text.Json;
using System.Text.Unicode;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

namespace constraints_summarizer_esapi_v15_5
{
    internal class CourseReferenceSheetDB
    {

        private static JsonSerializerOptions json_options = new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),  // for Japanese character
            ReadCommentHandling = JsonCommentHandling.Skip,         // allow commented JSON
            WriteIndented = true,       // for readability
        };

        private string db_file_path = "";

        string[] DB_INSTRUCTION =
        {
                "/*  \"患者ID - (Course, 制約ファイルの参照元)\" を対応させるデータベース",
                "id: 患者ID",
                "course: Course名。あるCourse毎に制約ファイルを紐づける。",
                "file: 制約ファイルの参照元",
                "！！　スクリプトの動作が重くなってきたら、以下のデータを削除",
                "もしくはファイルを別名保存してください。紐づけが初期化されます。　！！*/"
        };

        private PatientsLinkSet link_set = null;

        public CourseReferenceSheetDB(string db_path)
        {
            db_file_path = db_path;
            CreateDBFileIfNotExist();
        }


        public static string ToJson(Object link_set)
        {
            string ret;

            try
            {
                ret = JsonSerializer.Serialize(link_set, json_options);
            }
            catch (JsonException e)
            {
                ret = null;
            }

            return ret;
        }

        public static PatientsLinkSet JsonFileToLinkSet(string json_path)
        {
            PatientsLinkSet ret = null;
            string json = File.ReadAllText(json_path);

            if (String.IsNullOrEmpty(json))
            {
                ret = null;
            }
            try
            {
                ret = JsonSerializer.Deserialize<PatientsLinkSet>(json, json_options);
            }
            catch (JsonException e)
            {
                ret = null;
            }

            return ret;
        }

        public void CreateDBFileIfNotExist()
        {
            if (!File.Exists(db_file_path))
            {
                File.AppendAllLines(db_file_path, DB_INSTRUCTION);
//                MessageBox.Show(string.Format("DataBase file {0} is created.", db_file_path));
                MessageBox.Show(string.Format("データベースファイル '{0}' を作成しました。", db_file_path));
            }
            else
            {
//            MessageBox.Show("DB file is found");
            }

            return;
        }


        private CourseSheetLink GetLinkFromDB(string id, string course)
        {
            List<CourseSheetLink> links;
            CourseSheetLink ret = null;

            if (!File.Exists(db_file_path))
            {
                CreateDBFileIfNotExist();
                ret = null;
            }
            else
            {
                link_set = JsonFileToLinkSet(db_file_path);
                if (link_set == null)
                {
                    ret = null;
                }
                else if (link_set.Patients.TryGetValue(id, out links))
                {
                    CourseSheetLink link = links.FirstOrDefault(x => x.Course == course);
                    if (link == null)
                    {
                        ret = null;
                    }
                    else
                    {
                        ret = link;
                    }
                }
                else
                {
                    ret = null;
                }
            }

            return ret;
        }

        private CourseSheetLink AddLinkToDB(string id, string course, string file_path, bool overwrite = false)
        {
            //var link = new CourseSheetLink(course, file_path);
            var link = CourseSheetLinkBuilder(course, file_path);

            // if link_set is not exist, generate it
            if (link_set == null)
            {
                link_set = new PatientsLinkSet();
            }
            else {; }

            // register patients to link_set (link_set should not be null here)
            List<CourseSheetLink> links;
            if (link_set.Patients.TryGetValue(id, out links))
            {
                if (overwrite == true)      // for modifing consraints sheet PATH
                {
                    links.RemoveAll(x => x.Course == course);
                }
                else {; }
                links.Add(link);
            }
            else
            {
                links = new List<CourseSheetLink>() { link };
            }
            link_set.Patients[id] = links;


            var json = ToJson(link_set);
            File.WriteAllLines(db_file_path, DB_INSTRUCTION);
            File.AppendAllText(db_file_path, json);

            return link;
        }

        public string GetReferenceFilePath(string id, string course, string ref_path = null)
        {
            string file_path = null;

            // get file path
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (ref_path != null)
            {
                openFileDialog.InitialDirectory = ref_path;
            }
            else { }
            openFileDialog.Title = "Select a Constraints Reference File";
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls|All Files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                file_path = openFileDialog.FileName;
            }
            else {; }

            return file_path;
        }

        public CourseSheetLink GetOrAddLink(string id, string course, string ref_path)
        {
            CourseSheetLink ret = GetLinkFromDB(id, course);

            if (ret == null)
            {
                string file_path = GetReferenceFilePath(id, course, ref_path);
                ret = AddLinkToDB(id, course, file_path);
            }
            else {; }

            return ret;
        }

        public CourseSheetLink ModifyLink(string id, string course, string ref_path)
        {
            CourseSheetLink ret = null;
            string file_path = GetReferenceFilePath(id, course, ref_path);
            ret = AddLinkToDB(id, course, file_path, true);

            return ret;
        }

        private CourseSheetLink CourseSheetLinkBuilder(string course, string file)
        {
            var link = new CourseSheetLink();
            link.Course = course;
            link.FilePath = file;

            return link;
        }

    }

    public class CourseSheetLink
    {

        [JsonPropertyName("course")]
        public string Course { get; set; }

        [JsonPropertyName("file")]
        public string FilePath { get; set; }

        // if custom constructor is implemented, system.text.json serdes is not working properly.
        //public CourseSheetLink(string course, string file_path)
        //{
        //    Course = course;
        //    FilePath = file_path;
        //}
    }

    public class PatientsLinkSet
    {
        [JsonPropertyName("patients")]

        public Dictionary<string, List<CourseSheetLink>> Patients { get; set; } = new Dictionary<string, List<CourseSheetLink>>();
    }

}
