// Title: Find and replace currency‑formatted numeric cells in a named range with General format using Aspose.Cells for .NET
// AI Prompts: Search a workbook for numeric cells whose custom number format contains currency symbols within a specific named range and set their style to General with Aspose.Cells. | Write a C# method that takes a Workbook, a named range name, and a target format, then updates any currency‑styled numeric cells to the target format. | Add logging to capture the addresses of cells whose number format was changed from currency to General while processing a named range.
// Common Searches: Aspose.Cells .NET change currency number format to General in a defined name range | C# iterate over cells in a named range and remove currency formatting | How to detect currency symbols in cell style using Aspose.Cells | Convert formatted currency cells to plain numbers in Excel with Aspose.Cells for .NET | Update number format of numeric cells in a specific named range using Aspose.Cells API
// Tags: currency number format conversion Aspose.Cells .NET | named range cell style modification | detect currency symbols in cell format | set cell number format to General programmatically | iterate over defined name range Aspose.Cells | numeric cell type check Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, retrieves a named range, scans each numeric cell for a custom format containing currency symbols, changes those cells' style to the General format, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string namedRangeName = "MyNamedRange";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range by its defined name
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(namedRangeName);
            if (range == null)
            {
                Console.WriteLine("Named range not found.");
                return;
            }

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                // Process only numeric cells
                if (cell.Type == CellValueType.IsNumeric)
                {
                    // Get the cell's style to inspect its number format
                    Style style = cell.GetStyle();

                    // Use the custom number format string (if any)
                    string format = style.Custom;
                    if (!string.IsNullOrEmpty(format) &&
                        (format.Contains("$") || format.Contains("€") || format.Contains("£") ||
                         format.Contains("¥") || format.Contains("₹")))
                    {
                        // Change format to General (plain number)
                        style.Custom = "General";
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
