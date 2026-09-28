// Title: Compare initialized cell count before and after removing formatting‑only cells with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that counts cells containing data, clears styles from empty cells, then outputs the before and after counts. | Show how to detect empty cells with custom formatting, reset them to the default style, and calculate the reduction in initialized cells.
// Common Searches: how to count data cells in an Excel worksheet using Aspose.Cells C# | remove only formatting from empty cells Aspose.Cells .NET | compare cell count before and after style cleanup Aspose.Cells | C# Aspose.Cells reduce workbook size by clearing formatting‑only cells | measure initialized cells reduction after resetting cell styles with Aspose.Cells
// Tags: initialized cell count Aspose.Cells | remove formatting‑only cells Aspose.Cells | reset empty cell style Aspose.Cells | worksheet data cell enumeration C# | optimize workbook by clearing empty cell formats Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, counts cells that contain actual data, identifies empty cells that have custom formatting, resets those cells to the default style, recounts the initialized cells, prints the before/after counts and the reduction, and saves the cleaned workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Count initialized cells before removing formatting‑only cells
            int beforeCount = 0;
            foreach (Cell cell in cells)
            {
                // Consider a cell initialized if it contains a non‑null, non‑empty value
                if (cell.Value != null && !(cell.Value is string s && string.IsNullOrEmpty(s)))
                {
                    beforeCount++;
                }
            }

            // Remove cells that contain only formatting (no data)
            // Aspose.Cells may not expose RemoveFormattingOnlyCells in older versions,
            // so we clear the style of empty cells manually.
            Style defaultStyle = workbook.CreateStyle(); // default (no formatting) style
            List<Cell> cellsToClear = new List<Cell>();

            foreach (Cell cell in cells)
            {
                // Identify cells with no data but with custom formatting
                bool hasData = cell.Value != null && !(cell.Value is string str && string.IsNullOrEmpty(str));
                bool hasCustomStyle = cell.GetStyle() != null && !cell.GetStyle().Equals(defaultStyle);
                if (!hasData && hasCustomStyle)
                {
                    cellsToClear.Add(cell);
                }
            }

            foreach (Cell cell in cellsToClear)
            {
                cell.SetStyle(defaultStyle);
            }

            // Count initialized cells after the removal
            int afterCount = 0;
            foreach (Cell cell in cells)
            {
                if (cell.Value != null && !(cell.Value is string s && string.IsNullOrEmpty(s)))
                {
                    afterCount++;
                }
            }

            // Display the comparison results
            Console.WriteLine($"Initialized cells before removal: {beforeCount}");
            Console.WriteLine($"Initialized cells after removal:  {afterCount}");
            Console.WriteLine($"Reduction in initialized cells:   {beforeCount - afterCount}");

            // Save the modified workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
