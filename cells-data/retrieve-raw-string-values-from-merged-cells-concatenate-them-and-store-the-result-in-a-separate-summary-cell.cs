// Title: How to read raw values from merged cells and concatenate them into a summary cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get values from merged cells and combine into one cell | Aspose.Cells iterate merged areas and concatenate text in .NET | Read raw string from merged cell range and write summary cell using Aspose.Cells | How to concatenate merged cell contents into a single cell with Aspose.Cells C# | Aspose.Cells extract merged cell values without formulas C#
// Tags: merged cells value extraction Aspose.Cells | concatenate merged area strings C# | write summary cell PutValue Aspose.Cells | Aspose.Cells GetMergedAreas example | process merged cells workbook .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The code loads a workbook, iterates over each merged area on the first worksheet, reads the raw string from the top‑left cell of each merged region, concatenates these strings, writes the combined result into cell B1 using PutValue, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Load the workbook (existing or new) with safety checks.
        try
        {
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }
        }
        catch (CellsException ex)
        {
            // Handles password‑protected files or other loading issues.
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
            return;
        }

        // Process merged cells and save the result.
        try
        {
            Worksheet sheet = workbook.Worksheets[0];
            StringBuilder concatenated = new StringBuilder();

            // Use the recommended API to get merged areas.
            foreach (CellArea mergedArea in sheet.Cells.GetMergedAreas())
            {
                Cell cell = sheet.Cells[mergedArea.StartRow, mergedArea.StartColumn];
                string rawValue = cell.Value?.ToString() ?? string.Empty;
                concatenated.Append(rawValue);
            }

            // Write concatenated string to cell B1.
            sheet.Cells["B1"].PutValue(concatenated.ToString());

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
