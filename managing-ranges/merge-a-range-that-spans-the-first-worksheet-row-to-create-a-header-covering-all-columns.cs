// Title: Merge the first worksheet row across all columns to create a centered header using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge first worksheet row across all columns for header | Create a spanning header row in Excel with dynamic column count using Aspose.Cells | Apply centered bold style to a merged header cell in Aspose.Cells .NET | Determine number of columns to merge when worksheet has no data in Aspose.Cells | Save Excel file with merged header using Aspose.Cells C#
// Tags: merge row across columns Aspose.Cells | dynamic column span Aspose.Cells | centered bold header style Aspose.Cells | maxcolumn usage Aspose.Cells | save workbook merged header .NET

using Aspose.Cells;
using System;

// The example creates a new workbook, calculates the total columns to span (using the worksheet's MaxColumn or a default of 10), merges the entire first row into a single cell, inserts "Report Header", applies a centered bold style with a larger font, and saves the workbook as MergedHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the total number of columns to span.
        // If the sheet already has data, use the last used column; otherwise default to 10 columns.
        int maxColumn = sheet.Cells.MaxColumn; // zero‑based index of the last column with data
        int totalColumns = (maxColumn > 0) ? maxColumn + 1 : 10; // ensure at least 10 columns

        // Merge the first row (row index 0) across all columns to create a header.
        // Parameters: startRow, startColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 1, totalColumns);

        // Set header text in the merged cell (A1)
        sheet.Cells["A1"].PutValue("Report Header");

        // Apply simple styling to the header (centered, bold, larger font)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Size = 14;
        sheet.Cells["A1"].SetStyle(headerStyle);

        // Save the workbook to a file
        workbook.Save("MergedHeader.xlsx");
    }
}
