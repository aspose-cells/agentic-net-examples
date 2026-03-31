using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Microsoft.VisualBasic.FileIO; // Requires Microsoft.VisualBasic reference

namespace AsposeCellsVbaCsvDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths (adjust as needed)
            string inputPath = "InputWorkbook.xlsm";      // Existing macro‑enabled workbook
            string exportCsvPath = "ExportedModules.csv"; // CSV file to export VBA modules
            string importCsvPath = "ImportedModules.csv"; // CSV file to import VBA modules
            string outputPath = "UpdatedWorkbook.xlsm";   // Workbook after modifications

            // Ensure input workbook exists; create a simple one if missing
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbook(inputPath);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure a VBA project exists
            if (workbook.VbaProject == null)
            {
                // Save as macro‑enabled to embed a VBA project container
                workbook.Save(inputPath, SaveFormat.Xlsm);
                workbook = new Workbook(inputPath);
            }

            VbaProject vbaProject = workbook.VbaProject;

            // Export all VBA modules to a CSV file
            using (StreamWriter writer = new StreamWriter(exportCsvPath, false, Encoding.UTF8))
            {
                writer.WriteLine("ModuleName,ModuleCode");
                foreach (VbaModule module in vbaProject.Modules)
                {
                    string escapedCode = EscapeCsvField(module.Codes);
                    string escapedName = EscapeCsvField(module.Name);
                    writer.WriteLine($"{escapedName},{escapedCode}");
                }
            }

            // Import VBA modules from a CSV file and update the project
            if (File.Exists(importCsvPath))
            {
                using (TextFieldParser parser = new TextFieldParser(importCsvPath, Encoding.UTF8))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    bool firstLine = true;
                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (fields == null || fields.Length < 2)
                            continue;

                        if (firstLine)
                        {
                            firstLine = false;
                            if (fields.Length >= 2 && fields[0] == "ModuleName")
                                continue; // skip header
                        }

                        string moduleName = fields[0];
                        string moduleCode = fields[1];

                        // Unescape double quotes
                        moduleCode = moduleCode.Replace("\"\"", "\"");

                        VbaModule existingModule = null;
                        try
                        {
                            existingModule = vbaProject.Modules[moduleName];
                        }
                        catch { }

                        if (existingModule != null)
                        {
                            existingModule.Codes = moduleCode;
                        }
                        else
                        {
                            int newIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, moduleName);
                            VbaModule newModule = vbaProject.Modules[newIndex];
                            newModule.Codes = moduleCode;
                        }
                    }
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath, SaveFormat.Xlsm);
        }

        // Helper to escape a CSV field according to RFC 4180
        private static string EscapeCsvField(string field)
        {
            if (field.Contains("\"") || field.Contains(",") || field.Contains("\r") || field.Contains("\n"))
            {
                string escaped = field.Replace("\"", "\"\"");
                return $"\"{escaped}\"";
            }
            return field;
        }

        // Creates a minimal macro‑enabled workbook with a sample VBA module
        private static void CreateSampleWorkbook(string path)
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Name = "Sheet1";

            // Save as macro‑enabled to embed a VBA project container
            wb.Save(path, SaveFormat.Xlsm);

            // Re‑open to add VBA module
            Workbook macroWb = new Workbook(path);
            VbaProject proj = macroWb.VbaProject;
            int idx = proj.Modules.Add(VbaModuleType.Procedural, "SampleModule");
            VbaModule mod = proj.Modules[idx];
            mod.Codes = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

            macroWb.Save(path, SaveFormat.Xlsm);
        }
    }
}