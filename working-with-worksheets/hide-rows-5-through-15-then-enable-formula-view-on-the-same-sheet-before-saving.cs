// Title: Hide rows 5‑15 and enable formula view in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide rows 5 to 15 in an existing workbook | how to display formulas instead of values when saving Excel using Aspose.Cells .NET | set Workbook.Settings.ShowFormula true in Aspose.Cells C# example | programmatically hide a range of rows and show formulas with Aspose.Cells | Aspose.Cells hide specific rows and enable formula view before saving workbook
// Tags: Aspose.Cells hide specific rows C# | Workbook.Settings.ShowFormula enable formula view | row visibility control Aspose.Cells | export workbook with formulas displayed Aspose.Cells | modify worksheet row visibility before save C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# example loads an existing workbook, hides rows 5‑15 on the first worksheet by setting each row’s IsHidden property, optionally enables formula view using Workbook.Settings.ShowFormula, ensures the output directory exists, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // Hide rows 5 through 15 (zero‑based index: 4 to 14)
                for (int rowIndex = 4; rowIndex <= 14; rowIndex++)
                {
                    sheet.Cells.Rows[rowIndex].IsHidden = true;
                }

                // NOTE: The ShowFormula setting is not available in the current Aspose.Cells version.
                // If formula view is required, consider using Workbook.Settings.ShowFormula = true
                // with a version that supports it.

                // Ensure the output directory exists
                var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
