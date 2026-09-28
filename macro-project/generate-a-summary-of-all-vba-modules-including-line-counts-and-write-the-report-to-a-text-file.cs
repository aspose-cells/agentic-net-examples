// Title: Generate a VBA modules summary with line counts using Aspose.Cells in C# and save it to a text file
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates through each VBA module, counts the lines of code, and writes the module name and line count to a .txt report. | Enhance the utility to add a section that shows the total number of VBA modules and the cumulative line count in the generated summary file. | Modify the program to accept the input workbook path and the output report file path as command‑line arguments.
// Common Searches: how to extract VBA module names and line counts from an xlsx file using Aspose.Cells C# | C# program to list VBA modules in an Excel workbook and export a summary to a text file | Aspose.Cells count lines of code in each VBA module | generate VBA project report with Aspose.Cells in .NET | save VBA module statistics to a file using C#
// Tags: Aspose.Cells extract VBA modules | VBA module line count Aspose.Cells | export VBA summary to text file C# | enumerate VBA modules in Excel workbook | C# generate VBA project report

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The program loads an Excel workbook, checks for a VBA project, counts the lines in each VBA module, builds a summary report with module names and line counts, and writes the report to a text file.
class VbaSummaryGenerator
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare a StringBuilder to collect the report
            StringBuilder report = new StringBuilder();

            // Access the VBA project if it exists
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null || vbaProject.Modules.Count == 0)
            {
                report.AppendLine("No VBA project or modules found in the workbook.");
            }
            else
            {
                report.AppendLine("VBA Modules Summary:");
                report.AppendLine("--------------------");

                // Iterate through each VBA module
                foreach (VbaModule module in vbaProject.Modules)
                {
                    // Get the source code of the module
                    string code = module.Codes ?? string.Empty;

                    // Count lines (handles both Windows and Unix line endings)
                    int lineCount = 0;
                    if (!string.IsNullOrEmpty(code))
                    {
                        string[] lines = code.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                        lineCount = lines.Length;
                    }

                    // Append module information to the report
                    report.AppendLine($"Module Name : {module.Name}");
                    report.AppendLine($"Lines       : {lineCount}");
                    report.AppendLine();
                }
            }

            // Output report path
            string outputPath = "VbaModulesReport.txt";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the report to a text file
            File.WriteAllText(outputPath, report.ToString());

            Console.WriteLine($"VBA summary report generated at: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
