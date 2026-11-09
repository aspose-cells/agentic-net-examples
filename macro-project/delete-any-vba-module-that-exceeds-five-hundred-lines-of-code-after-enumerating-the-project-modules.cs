// Title: Remove VBA modules longer than 500 lines from an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete VBA modules that exceed 500 lines | how to filter and remove large VBA macros from an .xlsx file with Aspose.Cells | enumerate VBA project modules and purge those over a line count limit using Aspose.Cells .NET | example code to drop VBA modules larger than 500 lines in a workbook with Aspose.Cells | Aspose.Cells .NET remove VBA modules based on source code length
// Tags: remove large VBA modules Aspose.Cells C# | filter VBA modules by line count Aspose.Cells | delete VBA macros exceeding 500 lines .NET | enumerate Excel VBA project modules Aspose.Cells | save workbook after VBA module removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, checks for a VBA project, enumerates its modules, counts the source code lines of each module, removes any module containing more than 500 lines, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Check for VBA project and modules
            if (workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0)
            {
                var modulesToRemove = new System.Collections.Generic.List<dynamic>();

                // Identify modules with more than 500 lines
                foreach (dynamic vbaModule in workbook.VbaProject.Modules)
                {
                    string source = vbaModule.SourceCode ?? string.Empty;
                    int lineCount = source.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).Length;

                    if (lineCount > 500)
                    {
                        modulesToRemove.Add(vbaModule);
                    }
                }

                // Remove the identified modules
                foreach (dynamic module in modulesToRemove)
                {
                    workbook.VbaProject.Modules.Remove(module);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
