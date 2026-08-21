// Title: Aspose.Cells .NET – Create a VBA Module Summary with Line Counts and Export to Text
// Description: C# code that loads a macro‑enabled Excel workbook with Aspose.Cells, detects a VBA project, enumerates each VbaModule, counts its code lines, formats a concise report (module name, type, line count) and writes the result to a text file. Includes file‑existence checks and robust exception handling.
// Keywords: Aspose.Cells VBA module summary | C# count VBA lines | export VBA report to text | macro‑enabled workbook analysis | VbaProject module enumeration | .NET Excel VBA statistics | generate VBA code metrics
// Common Searches: list VBA modules and line counts using Aspose.Cells .NET | how to export VBA module details to a text file in C# | count lines of code in each Excel VBA module | Aspose.Cells generate VBA project report | C# extract VBA module information from workbook
// Developer Intent: Extract every VBA module from a macro‑enabled workbook, compute the number of code lines per module, and produce a readable text summary.
// Use Cases: Audit macro complexity by obtaining line‑count statistics for all VBA modules. | Document VBA project structure for version‑control, code reviews, or compliance reporting. | Automate batch processing to generate module metrics across multiple workbooks on a schedule.
// AI Prompts: Write a C# method that uses Aspose.Cells to return a dictionary of VBA module names and their line counts. | Suggest improvements for logging and error handling in the VBA summary generator, including writing errors to a separate log file. | Show how to modify the report to include module descriptions and export the data as CSV instead of plain text.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // C# code that loads a macro‑enabled Excel workbook with Aspose.Cells, detects a VBA project, enumerates each VbaModule, counts its code lines, formats a concise report (module name, type, line count) and writes the result to a text file. Includes file‑existence checks and robust exception handling.
    public class VbaModulesSummaryGenerator
    {
        /// <param name="workbookPath">Path to the Excel workbook (must be macro‑enabled if it contains VBA).</param>
        /// <param name="reportPath">Path where the summary text file will be saved.</param>
        public static void Run(string workbookPath, string reportPath)
        {
            try
            {
                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                    throw new FileNotFoundException("Workbook file not found.", workbookPath);

                // Load the workbook from the given file path
                Workbook workbook = new Workbook(workbookPath);

                // Prepare a StringBuilder to collect the report content
                StringBuilder reportBuilder = new StringBuilder();

                // Check whether the workbook contains a VBA project
                if (workbook.VbaProject == null)
                {
                    reportBuilder.AppendLine("No VBA project found in the workbook.");
                }
                else
                {
                    VbaModuleCollection modules = workbook.VbaProject.Modules;

                    // If there are no modules, note that in the report
                    if (modules.Count == 0)
                    {
                        reportBuilder.AppendLine("VBA project exists but contains no modules.");
                    }
                    else
                    {
                        reportBuilder.AppendLine($"Total VBA modules: {modules.Count}");
                        reportBuilder.AppendLine();

                        // Iterate through each module and compute line count
                        for (int i = 0; i < modules.Count; i++)
                        {
                            VbaModule module = modules[i];
                            string moduleName = module.Name ?? $"Module_{i}";
                            string code = module.Codes ?? string.Empty;

                            // Split on both CRLF and LF to count lines accurately
                            int lineCount = code.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Length;

                            reportBuilder.AppendLine($"Module Name : {moduleName}");
                            reportBuilder.AppendLine($"Module Type : {module.Type}");
                            reportBuilder.AppendLine($"Line Count  : {lineCount}");
                            reportBuilder.AppendLine(new string('-', 40));
                        }
                    }
                }

                // Write the assembled report to the specified text file
                File.WriteAllText(reportPath, reportBuilder.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during report generation: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <workbookPath> <reportPath>");
                return;
            }

            string workbookPath = args[0];
            string reportPath = args[1];

            try
            {
                VbaModulesSummaryGenerator.Run(workbookPath, reportPath);
                Console.WriteLine($"Report successfully generated at: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
            }
        }
    }
}
