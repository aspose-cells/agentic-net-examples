using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSummary
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file (must be a macro‑enabled workbook, e.g., .xlsm)
            string inputPath = "input.xlsm";

            // Output text file that will contain the summary
            string reportPath = "VbaModulesReport.txt";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project; if none exists, exit
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Get the collection of VBA modules
            VbaModuleCollection modules = vbaProject.Modules;

            // Prepare to write the report
            using (StreamWriter writer = new StreamWriter(reportPath))
            {
                writer.WriteLine($"VBA Modules Summary for '{Path.GetFileName(inputPath)}'");
                writer.WriteLine(new string('=', 50));

                // Iterate through each module
                for (int i = 0; i < modules.Count; i++)
                {
                    VbaModule module = modules[i];
                    string code = module.Codes ?? string.Empty;

                    // Count non‑empty lines in the module code
                    int lineCount = 0;
                    if (!string.IsNullOrEmpty(code))
                    {
                        // Split on both CR and LF to handle different line endings
                        string[] lines = code.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        lineCount = lines.Length;
                    }

                    // Write module details to the report
                    writer.WriteLine($"Module {i + 1}:");
                    writer.WriteLine($"  Name      : {module.Name}");
                    writer.WriteLine($"  Type      : {module.Type}");
                    writer.WriteLine($"  Line Count: {lineCount}");
                    writer.WriteLine();
                }

                writer.WriteLine("End of Summary");
            }

            Console.WriteLine($"VBA modules summary written to '{reportPath}'.");
        }
    }
}