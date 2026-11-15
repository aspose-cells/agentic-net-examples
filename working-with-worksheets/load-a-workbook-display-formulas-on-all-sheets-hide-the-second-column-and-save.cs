// Title: Load an Excel workbook, hide column B in every worksheet, and save the file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide column B in all worksheets | C# example to load an .xlsx, hide second column, and save using Aspose.Cells | How to hide a specific column across multiple sheets with Aspose.Cells .NET | Workaround for displaying formulas when ShowFormula is unavailable in Aspose.Cells | Ensure output folder exists before saving workbook with Aspose.Cells C#
// Tags: hide column index Aspose.Cells C# | iterate worksheets hide column Aspose.Cells | load workbook modify hide column save Aspose.Cells | formula display workaround Aspose.Cells .NET | create output directory Aspose.Cells save

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# program loads an existing Excel file using Aspose.Cells, iterates through each worksheet to hide column B (the second column), ensures the target output directory exists, and saves the modified workbook. It also notes that the ShowFormula property is not available in the current version and suggests a workaround for displaying formulas.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                var workbook = new Workbook(inputPath);

                // NOTE: The ShowFormula property is not available in this version of Aspose.Cells.
                // If displaying formulas is required, consider using alternative approaches such as
                // iterating cells and retrieving their Formula property.

                // Process each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Hide the second column (zero‑based index 1)
                    sheet.Cells.HideColumn(1);
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
