// Title: Identify formatting‑only initialized cells in an Aspose.Cells worksheet using a C# extension method
// AI Prompts: Write a C# extension method for Aspose.Cells that scans a worksheet's used range and returns true when a cell has a style different from the workbook's default style while its Value is null. | Refactor the extension to return a collection of cell addresses that contain only formatting (non‑default style, no value) instead of a simple boolean. | Create NUnit test cases in C# that validate the extension correctly flags formatting‑only cells and ignores cells that contain data.
// Common Searches: Aspose.Cells C# how to detect empty cells with custom formatting | C# extension method to find cells that have style but no value in an Excel worksheet | Check worksheet for formatting‑only initialized cells using Aspose.Cells API | Identify cells with non‑default style and no data in Aspose.Cells C# | Determine if a worksheet contains cells that are formatted but not populated in Aspose.Cells
// Tags: detect formatting‑only cells Aspose.Cells | worksheet style comparison default Aspose.Cells | extension method cell style analysis C# | maxdisplayrange empty cell formatting check | non‑default cell style detection Excel C#

using System;
using System.IO;
using Aspose.Cells;

// Provides a C# WorksheetExtensions class with a HasFormattingOnlyInitializedCells extension method that iterates over the worksheet's MaxDisplayRange, compares each cell's style to the workbook's default style, and returns true if any cell lacks a value yet has a custom style, enabling detection of formatting‑only initialized cells.
public static class WorksheetExtensions
{
    /// <param name="sheet">The worksheet to inspect.</param>
    /// <returns>True when a formatting‑only initialized cell is found; otherwise false.</returns>
    public static bool HasFormattingOnlyInitializedCells(this Worksheet sheet)
    {
        // Get the default style of the workbook for comparison.
        Style defaultStyle = sheet.Workbook.DefaultStyle;

        // Determine the used range of the worksheet.
        // MaxDisplayRange covers the area that has been accessed or formatted.
        var range = sheet.Cells.MaxDisplayRange;
        int firstRow = range.FirstRow;
        int firstColumn = range.FirstColumn;
        int rowCount = range.RowCount;
        int columnCount = range.ColumnCount;

        // Iterate through each cell in the used range.
        for (int i = firstRow; i < firstRow + rowCount; i++)
        {
            for (int j = firstColumn; j < firstColumn + columnCount; j++)
            {
                Cell cell = sheet.Cells[i, j];

                // Skip cells that already contain a value.
                if (cell.Value != null)
                    continue;

                // Retrieve the cell's style.
                Style cellStyle = cell.GetStyle();

                // If the style differs from the default, the cell is formatting‑only.
                if (!cellStyle.Equals(defaultStyle))
                    return true;
            }
        }

        // No formatting‑only cells were found.
        return false;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string filePath = "sample.xlsx";
            Workbook workbook;

            // Load existing workbook if present; otherwise create a new one.
            if (File.Exists(filePath))
            {
                workbook = new Workbook(filePath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Use the extension method to check for formatting‑only initialized cells.
            bool hasFormattingOnly = sheet.HasFormattingOnlyInitializedCells();

            Console.WriteLine($"Formatting‑only initialized cells found: {hasFormattingOnly}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
