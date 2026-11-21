// Title: How to freeze the top row after AutoFitRows in Aspose.Cells for .NET to keep row heights fixed while scrolling
// Common Searches: Aspose.Cells C# freeze first row after calling AutoFitRows | keep row heights static while scrolling Excel worksheet using Aspose.Cells | apply FreezePanes after AutoFitRows in .NET workbook example | C# code to auto-fit rows then freeze top pane with Aspose.Cells | prevent row height changes on scroll after AutoFitRows Aspose.Cells
// Tags: AutoFitRows then FreezePanes in Aspose.Cells | freeze first row after row autofit C# | lock row heights during scroll Aspose.Cells | worksheet FreezePanes usage .NET | Excel workbook row height preservation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, calls AutoFitRows on the first worksheet to adjust all row heights, then freezes the top row using FreezePanes(1,0,1,0) so those heights remain constant while scrolling, and finally saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Auto‑fit all rows in the worksheet
                sheet.AutoFitRows();

                // Freeze the top row (row index 0)
                // Parameters: firstRow, firstColumn, totalRows, totalColumns
                sheet.FreezePanes(1, 0, 1, 0);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
