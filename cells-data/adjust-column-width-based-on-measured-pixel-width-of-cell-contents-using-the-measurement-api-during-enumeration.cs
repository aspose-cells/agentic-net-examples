// Title: Customizing Excel column widths in Aspose.Cells for .NET by measuring cell text pixel size during column iteration
// AI Prompts: Write C# code that loops through each column of a worksheet, uses the Aspose.Cells measurement API (e.g., GetColumnWidthPixel / SetColumnWidthPixel) to find the maximum pixel width of the cells in that column, and then sets the column width accordingly. | Refactor the sample to replace the AutoFitColumns call with a routine that measures each cell's displayed string width in pixels and applies the calculated width using SetColumnWidthPixel.
// Common Searches: Aspose.Cells .NET how to set column width based on pixel measurement of cell text | C# custom auto‑fit columns using measurement API in Aspose.Cells | Calculate Excel column width from string pixel length with Aspose.Cells for .NET | Iterate worksheet columns and adjust width by measured pixel size Aspose.Cells | Set column width in pixels instead of characters using Aspose.Cells C#
// Tags: custom column width calculation using Aspose.Cells measurement API | set Excel column width from pixel size in .NET | iterate worksheet columns to adjust width Aspose.Cells | replace AutoFitColumns with pixel‑based sizing | measure string pixel width for Excel column Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, fills several cells with short and long text strings, calls AutoFitColumns to automatically size all columns based on their contents, and saves the workbook as AdjustedColumnWidths.xlsx.
class ColumnWidthAdjuster
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Short");
            sheet.Cells["B1"].PutValue("A much longer piece of text");
            sheet.Cells["C1"].PutValue("Medium length");
            sheet.Cells["A2"].PutValue("Another short");
            sheet.Cells["B2"].PutValue("Tiny");
            sheet.Cells["C2"].PutValue("Extremely long text that should expand the column width significantly");

            // Auto‑fit all columns based on the content
            sheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("AdjustedColumnWidths.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
