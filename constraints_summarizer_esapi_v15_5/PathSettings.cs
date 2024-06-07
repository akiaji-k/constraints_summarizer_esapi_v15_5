using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace constraints_summarizer_esapi_v15_5
{
    internal class PathSettings
    {

        public string DbFilePath { get; private set; } = "";
        public string ConstraintsReferenceDirectory { get; private set; } = "";
        public string OutputDirectory { get; private set; } = "";

        public PathSettings(in string file_path)
        {
            try
            {
                var parser = new TextFieldParser(file_path);
                parser.Delimiters = new string[] { "," };
                string foo = "";
                while (!parser.EndOfData)
                {
                    var fields = parser.ReadFields();

                    switch (fields[0])
                    {
                        case "DBFilePath":
                            if (fields.Length != 2)
                            {
                                foo += "DBFilePath has invalid filed.\n";
                            }
                            else if (String.IsNullOrEmpty(fields[1]))
                            {
                                foo += "DBFilePath has empty filed.\n";
                            }
                            else
                            {
                                DbFilePath = fields[1];

                                if (!File.Exists(fields[1]))
                                {
                                    foo += string.Format("Database file {0} is not exist.\n", fields[1]);
                                }
                                else {; }

                            }
                            break;

                        case "ConstraintsReferenceDirectory":
                            if (fields.Length != 2)
                            {
                                foo += "ConstraintsReferenceDirectory has invalid filed.\n";
                            }
                            else if (String.IsNullOrEmpty(fields[1]))
                            {
                                foo += "ConstraintsReferenceDirectory has empty filed.\n";
                            }
                            else if (!Directory.Exists(fields[1]))
                            {
                                foo += string.Format("ConstraintsReferenceDirectory {0} is not exist.\n", fields[1]);
                            }
                            else
                            {
                                ConstraintsReferenceDirectory = fields[1];
                            }
                            break;
                        case "OutputDirectory":
                            if (fields.Length != 2)
                            {
                                foo += "OutputDirectory has invalid filed.\n";
                            }
                            else if (String.IsNullOrEmpty(fields[1]))
                            {
                                foo += "OutputDirectory has empty filed.\n";
                            }
                            else if (!Directory.Exists(fields[1]))
                            {
                                foo += string.Format("OutputDirectory {0} is not exist.\n", fields[1]);
                            }
                            else
                            {
                                OutputDirectory = fields[1];
                            }
                            break;

                        default:
                            break;

                    }
                }

                if (foo != "")
                {
                    MessageBox.Show(foo);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return;
        }
    }
}
