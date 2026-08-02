// Title: Count non‑empty cells in a worksheet's MaxDisplayRange with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, retrieves Cells.MaxDisplayRange, iterates through the range, counts cells whose Value is not null, handles an empty sheet, and saves the file.
// Keywords: Aspose.Cells | MaxDisplayRange | count non‑empty cells | C# example | iterate cells | worksheet density | empty worksheet handling | Aspose.Cells for .NET | cell value null check | Excel automation
// Common Searches: C# count non‑empty cells Aspose.Cells | How to use Cells.MaxDisplayRange | Get number of filled cells in Excel with Aspose | Iterate over MaxDisplayRange C# | Handle empty worksheet Aspose.Cells | Count populated cells in worksheet range .NET
// Developer Intent: Determine how many cells contain data inside the worksheet's maximum display range.
// Use Cases: Validate that a generated report contains the expected number of data rows by counting populated cells in MaxDisplayRange. | Skip processing of empty rows or columns when exporting data after first measuring cell density in the display range. | Provide a quick worksheet summary (e.g., fill‑rate) before saving or publishing the Excel file.
// AI Prompts: Write a C# method using Aspose.Cells that returns the count of non‑null cells in a worksheet's MaxDisplayRange. | Show how to log the address of each non‑empty cell while iterating over MaxDisplayRange in Aspose.Cells for .NET. | Explain safe handling of a null MaxDisplayRange when counting cells in an empty worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, retrieves Cells.MaxDisplayRange, iterates through the range, counts cells whose Value is not null, handles an empty sheet, and saves the file.
    public class CountNonEmptyCellsInMaxDisplayRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data (including empty cells)
                cells["A1"].PutValue("Header");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue(123);
                cells["C3"].PutValue("Text");
                // Leave some cells empty intentionally

                // Get the maximum display range that includes data, merged cells and shapes
                AsposeRange maxDisplayRange = cells.MaxDisplayRange;

                // If the worksheet is empty, MaxDisplayRange will be null
                if (maxDisplayRange == null)
                {
                    Console.WriteLine("The worksheet is empty. Non‑empty cell count: 0");
                    return;
                }

                // Iterate through all cells in the range and count those with a non‑null value
                int nonEmptyCount = 0;
                foreach (Cell cell in maxDisplayRange)
                {
                    if (cell.Value != null)
                    {
                        nonEmptyCount++;
                    }
                }

                Console.WriteLine($"Non‑empty cells in MaxDisplayRange: {nonEmptyCount}");

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                string outputPath = "CountNonEmptyCellsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            CountNonEmptyCellsInMaxDisplayRange.Run();
        }
    }
}
