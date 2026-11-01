// Title: How to relocate a PivotTable to a new cell range while preserving its layout using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move pivot table to a specific cell address | preserve pivot table layout when changing its position with Aspose.Cells | programmatically relocate an Excel pivot table using Aspose.Cells .NET | move pivot table to F31 using Aspose.Cells API | cut and paste pivot table range without losing formatting in C#
// Tags: Aspose.Cells PivotTable.Move usage | move pivot table cell range C# | preserve pivot structure Aspose.Cells | relocate pivot table in Excel workbook .NET | programmatic pivot table repositioning Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, checks for a PivotTable on the first worksheet, and uses the PivotTable.Move method to shift the pivot to cell F31 while keeping its structure intact. The modified workbook is then saved to a new file, with error handling for missing files and other exceptions.
class PivotTableCutPaste
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table on the sheet
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Define the destination top‑left cell for the move operation (F31)
            int destFirstRow = 30;        // Row 31 (zero‑based)
            int destFirstColumn = 5;      // Column F (zero‑based)

            // Move the pivot table to the new location
            pivotTable.Move(destFirstRow, destFirstColumn);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
