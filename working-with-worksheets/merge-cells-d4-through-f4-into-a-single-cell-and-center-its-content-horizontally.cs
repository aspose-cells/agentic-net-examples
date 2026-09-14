// Title: How to merge cells D4‑F4 and horizontally center the text using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to merge the range D4:F4 on the first worksheet and set its horizontal alignment to center. | Generate a complete Aspose.Cells example that creates a workbook, merges cells D4 through F4, applies a centered style, and saves the file.
// Common Searches: aspnet aspose.cells merge D4 to F4 and center text | c# aspose.cells merge cells and set horizontal alignment | how to merge a row of cells and center content with Aspose.Cells for .NET | aspose.cells example merging cells D4-F4 and applying style
// Tags: merge cell range D4-F4 Aspose.Cells | apply horizontal center alignment Aspose.Cells | create workbook and merge cells C# | set style for merged cells Aspose.Cells | save workbook with merged cells Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, merges cells D4 through F4 on the first worksheet, centers the merged cell's content horizontally, and saves the result as MergedCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells D4 through F4 (zero‑based indices: row 3, column 3)
        // Parameters: startRow, startColumn, totalRows, totalColumns
        sheet.Cells.Merge(3, 3, 1, 3);

        // Apply horizontal center alignment to the merged cell
        Style style = sheet.Cells["D4"].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        sheet.Cells["D4"].SetStyle(style);

        // Save the workbook (lifecycle: save)
        workbook.Save("MergedCell.xlsx");
    }
}
