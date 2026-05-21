using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace VbaProjectJsonReport
{
    // Represents a simplified VBA module information for JSON serialization
    public class VbaModuleInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Codes { get; set; }
    }

    // Represents the VBA project report structure
    public class VbaProjectReport
    {
        public string ProjectName { get; set; }
        public List<VbaModuleInfo> Modules { get; set; } = new List<VbaModuleInfo>();
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (macro-enabled). Adjust as needed.
            string inputPath = "input.xlsm";

            // Load the workbook; if it does not exist, create a new macro-enabled workbook and add a sample module.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Ensure the workbook is saved as macro-enabled later.
                // Add a sample VBA module so that the report contains data.
                int idx = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "SampleModule");
                VbaModule sampleModule = workbook.VbaProject.Modules[idx];
                sampleModule.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Build the report object
            VbaProjectReport report = new VbaProjectReport
            {
                ProjectName = vbaProject.Name
            };

            // Iterate through all modules and collect their details
            foreach (VbaModule module in vbaProject.Modules)
            {
                VbaModuleInfo info = new VbaModuleInfo
                {
                    Name = module.Name,
                    Type = module.Type.ToString(),
                    Codes = module.Codes ?? string.Empty
                };
                report.Modules.Add(info);
            }

            // Serialize the report to JSON with indentation for readability
            JsonSerializerOptions options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            string json = JsonSerializer.Serialize(report, options);

            // Define the output JSON file path
            string outputPath = "VbaProjectReport.json";

            // Write the JSON content to the file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"VBA project report has been generated at: {Path.GetFullPath(outputPath)}");

            // Save the workbook as macro-enabled if it was newly created
            if (!File.Exists(inputPath))
            {
                workbook.Save("GeneratedWorkbook.xlsm", SaveFormat.Xlsm);
            }
        }
    }
}