// Title: Merge the first worksheet row across multiple columns to create a centered header using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge first row cells across columns for header | Create a spanning header in Excel with Aspose.Cells .NET | How to set centered bold text in a merged header row using Aspose.Cells | Merge range A1:J1 and apply style with Aspose.Cells C# example | Save workbook after merging first row cells with Aspose.Cells for .NET
// Tags: merge cells first row Aspose.Cells | spanning header Excel .NET | centered bold header style Aspose.Cells | merge range A1:J1 Aspose.Cells | apply style to merged cell Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, merges the cells in the first row across the first ten columns (A‑J), inserts "Report Header" into the merged cell, applies centered bold styling, and saves the file as MergedHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define how many columns the header should span.
        // Here we assume a header covering the first 10 columns (A to J).
        int totalColumns = 10;

        // Merge the cells in the first row across the defined columns.
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 1, totalColumns);

        // Set the header text in the merged cell (top‑left cell of the range)
        sheet.Cells[0, 0].PutValue("Report Header");

        // Optionally, apply some styling to the header
        Style style = sheet.Cells[0, 0].GetStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;
        style.Font.Size = 14;
        sheet.Cells[0, 0].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("MergedHeader.xlsx");
    }
}
