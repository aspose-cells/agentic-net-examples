// Title: Create a new workbook, merge range A1:D1, apply centered alignment, and save as XLSX with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to create a workbook, merge cells A1 through D1, insert a title, apply both horizontal and vertical centering, and write the file to output.xlsx. | Demonstrate how to define a style with centered alignment and assign it to a merged range in Aspose.Cells for .NET.
// Common Searches: asp.net c# merge A1:D1 and center text using Aspose.Cells | how to apply vertical and horizontal alignment to a merged cell in Aspose.Cells | save a workbook with a centered title row as XLSX with Aspose.Cells .NET | Aspose.Cells create workbook, merge cells, set style, and export to Excel
// Tags: merge cells A1:D1 Aspose.Cells | centered style for merged range C# | create workbook and export XLSX Aspose.Cells | apply horizontal and vertical alignment Aspose.Cells

using Aspose.Cells;
using System;

// C# program that creates a new workbook, merges cells A1 through D1, writes a title, applies a style with both horizontal and vertical center alignment to the merged cell, and saves the result as output.xlsx using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells A1:D1 (row 0, column 0, 1 row, 4 columns)
        sheet.Cells.Merge(0, 0, 1, 4);

        // Optional: set a title text
        sheet.Cells["A1"].PutValue("Title");

        // Create a style with center alignment
        Style style = workbook.CreateStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;

        // Apply the style to the merged cell
        sheet.Cells["A1"].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
