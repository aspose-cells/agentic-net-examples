// Title: How to merge cells A1:F2 into a centered title block with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge range A1:F2 and center the text | Create a merged title cell that spans multiple rows and columns using Aspose.Cells | Apply bold 16‑point font to a merged Excel cell with Aspose.Cells .NET | Use Cells.Merge to combine rows and columns in a worksheet via Aspose.Cells API | Set horizontal and vertical alignment for a merged range in Aspose.Cells C#
// Tags: Cells.Merge method usage | merged title block Excel | centered text in merged range | bold 16pt font style Aspose.Cells | save workbook to XLSX file

using Aspose.Cells;
using System;

// The example creates a new workbook, merges cells A1 through F2 into a single title block, inserts "Report Title", centers the text both horizontally and vertically, applies bold 16‑point formatting, and saves the result as MergedTitle.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells from A1 to F2 (rows 0-1, columns 0-5)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 2, 6);

        // Set title text in the merged cell
        Cell titleCell = sheet.Cells["A1"];
        titleCell.PutValue("Report Title");

        // Apply formatting to the title
        Style style = titleCell.GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;
        style.Font.Size = 16;
        titleCell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("MergedTitle.xlsx");
    }
}
