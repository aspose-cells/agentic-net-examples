// Title: C# extension method to check for formatting‑only cells in an Aspose.Cells worksheet
// AI Prompts: Write a C# extension method for Aspose.Cells that returns true when a cell has a non‑default style but no value. | Generate code that iterates over a worksheet's MaxDisplayRange, skips any cell containing data, and detects styled empty cells. | Create a sample program that loads or creates a workbook, applies custom formatting to an empty cell, and uses the extension method to report formatting‑only cells.
// Common Searches: Aspose.Cells C# find cells with custom formatting but no data | How to detect empty styled cells in a worksheet using Aspose.Cells | C# check if worksheet contains formatting‑only cells with Aspose.Cells API | Identify cells that have style changes but no value in Aspose.Cells workbook | Use MaxDisplayRange to locate formatting‑only cells in Aspose.Cells C#
// Tags: detect formatting‑only cells Aspose.Cells | worksheet style comparison without values C# | MaxDisplayRange styled empty cell detection | extension method for Aspose.Cells worksheet analysis | default style comparison Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Provides a C# extension method (HasFormattingOnlyCells) that scans a worksheet's MaxDisplayRange, ignores cells with any value (including empty strings), compares each cell's style to the default style, and returns true as soon as it finds a cell that has custom formatting but no content. Includes a demo program that creates or loads a workbook, applies formatting to an empty cell, and prints the detection result.
public static class WorksheetExtensions
{
    /// <param name="worksheet">The worksheet to inspect.</param>
    /// <returns>True if a formatting‑only cell exists; otherwise, false.</returns>
    public static bool HasFormattingOnlyCells(this Worksheet worksheet)
    {
        // Get the area that actually contains data or formatting.
        var maxRange = worksheet.Cells.MaxDisplayRange;
        if (maxRange == null)
            return false; // No cells at all.

        // Capture the default style (the style of a newly created cell).
        Style defaultStyle = worksheet.Cells[0, 0].GetStyle();

        int firstRow = maxRange.FirstRow;
        int firstColumn = maxRange.FirstColumn;
        int totalRows = maxRange.RowCount;
        int totalColumns = maxRange.ColumnCount;

        for (int row = firstRow; row < firstRow + totalRows; row++)
        {
            for (int col = firstColumn; col < firstColumn + totalColumns; col++)
            {
                Cell cell = worksheet.Cells[row, col];

                // Skip cells that have a value (including empty string).
                if (cell.Value != null && !(cell.Value is string s && string.IsNullOrEmpty(s)))
                    continue;

                // Compare the cell's style with the default style.
                Style style = cell.GetStyle();
                if (!style.Equals(defaultStyle))
                {
                    // Found a cell with formatting but no value.
                    return true;
                }
            }
        }

        // No formatting‑only cells were found.
        return false;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            const string filePath = "sample.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(filePath))
            {
                workbook = new Workbook(filePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Create a formatting‑only cell for demonstration.
                Cell fmtCell = ws.Cells["B2"];
                fmtCell.PutValue(string.Empty);
                Style style = fmtCell.GetStyle();
                style.Font.IsBold = true;
                fmtCell.SetStyle(style);
                // Save the workbook for future runs.
                workbook.Save(filePath);
            }

            Worksheet worksheet = workbook.Worksheets[0];

            bool hasFormattingOnly = worksheet.HasFormattingOnlyCells();

            Console.WriteLine($"Worksheet contains formatting‑only cells: {hasFormattingOnly}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
