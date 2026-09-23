// Title: How to unmerge cells that span row 10 in an Excel file while preserving their values using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, finds every merged range intersecting row 10, copies the original top‑left cell value into all cells of each range, removes the merge, and saves the result. | Create a reusable method in C# that takes input and output file paths, detects merged areas covering a given row, duplicates the source value across the unmerged cells, and persists the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# unmerge cells that include a specific row and keep data | preserve cell values when removing merged cells in Excel using Aspose.Cells .NET | list merged areas that cover row 10 using Aspose.Cells | C# code to unmerge row 10 merged cells and copy original value to each cell
// Tags: unmerge cells in row 10 Aspose.Cells | copy original merged cell value to each cell .NET | retrieve merged areas worksheet C# | Aspose.Cells preserve values after unmerge | load and save workbook with merged cell handling Aspose

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, identifies all merged ranges that intersect row 10, copies the original top‑left cell value into every cell of each range, removes the merged areas, and saves the modified workbook, with file‑existence checks and error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Row 10 in zero‑based index is 9
            const int targetRow = 9;

            // Collect merged areas that intersect row 10
            List<CellArea> areasToUnmerge = new List<CellArea>();
            foreach (CellArea area in sheet.Cells.GetMergedAreas())
            {
                if (area.StartRow <= targetRow && area.EndRow >= targetRow)
                {
                    areasToUnmerge.Add(area);
                }
            }

            // Process each collected merged area
            foreach (CellArea area in areasToUnmerge)
            {
                // Preserve the original value (typically stored in the top‑left cell)
                object originalValue = sheet.Cells[area.StartRow, area.StartColumn].Value;

                // Fill every cell in the merged region with the preserved value
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        sheet.Cells[row, col].Value = originalValue;
                    }
                }

                // Unmerge the cells by removing the merged area from the collection
                sheet.Cells.MergedCells.Remove(area);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
