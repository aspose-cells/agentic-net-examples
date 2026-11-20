// Title: Merge cells F4:G5, apply a formula, and export the worksheet to CSV with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to merge a specific cell range and set a formula before saving as CSV | Aspose.Cells .NET merge cells F4:G5 and reference the merged cell in a formula | Export merged cells to CSV using Aspose.Cells in a C# console application | Set formula in a merged range with Aspose.Cells and generate CSV output | Load or create workbook, merge cells and save as CSV using Aspose.Cells C#
// Tags: merge cell block F4:G5 Aspose.Cells | assign formula to merged cell C# | export worksheet to CSV using Aspose.Cells | load existing workbook or create new Aspose.Cells | cell merging with formula for CSV output .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook or creates a new one, merges the range F4:G5 on the first worksheet, assigns the formula "=F4*2" to the merged cell, and saves the result as a CSV file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    // Load existing workbook (no password required)
                    var loadOptions = new LoadOptions();
                    workbook = new Workbook(inputPath, loadOptions);
                }
                else
                {
                    // Create a new workbook with a default worksheet
                    workbook = new Workbook();
                    // Optional: put a sample value in the target cell
                    workbook.Worksheets[0].Cells["F4"].PutValue(10);
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Merge cells F4:G5 (row 3, column 5) spanning 2 rows and 2 columns
                sheet.Cells.Merge(3, 5, 2, 2);

                // Set formula referencing the merged cell
                sheet.Cells["F4"].Formula = "=F4*2";

                // Save the worksheet as CSV
                const string outputPath = "output.csv";
                workbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
