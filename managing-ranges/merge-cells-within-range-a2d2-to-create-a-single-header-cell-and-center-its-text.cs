// Title: How to merge cells A2:D2 into a single centered header using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a row of cells and center the header text | Create a merged header spanning A2 to D2 with centered alignment in Excel using Aspose.Cells | Set horizontal and vertical alignment for a merged range in Aspose.Cells .NET | C# code to merge cells and apply StyleFlag for a centered header in a workbook | Aspose.Cells example for merging cells and centering text in a header row
// Tags: merge range A2:D2 Aspose.Cells | apply centered alignment merged cells | StyleFlag usage for merged cell styling | Excel header row merge C# Aspose.Cells | horizontal vertical text alignment Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, merges cells A2 through D2 on the first worksheet, inserts the text "Header", applies a style that centers the text both horizontally and vertically using StyleFlag, and saves the file as output.xlsx.
class MergeHeaderExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells A2:D2 to form a single header cell (range: row 1, column 0 to row 1, column 3)
        sheet.Cells.Merge(1, 0, 1, 4); // Merge 1 row, 4 columns starting at A2

        // Set the header text
        sheet.Cells["A2"].PutValue("Header");

        // Create a style to center the text horizontally and vertically
        Style style = workbook.CreateStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;

        // Apply the style to the merged cell
        StyleFlag flag = new StyleFlag();
        flag.All = true; // Apply all style attributes
        sheet.Cells["A2"].SetStyle(style, flag);

        // Save the workbook (lifecycle save rule)
        workbook.Save("output.xlsx");
    }
}
