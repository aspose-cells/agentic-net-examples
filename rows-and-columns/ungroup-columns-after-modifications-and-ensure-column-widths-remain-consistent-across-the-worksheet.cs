// Title: Ungroup Columns and Keep Original Widths with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, records the widths of a column range, groups the columns, temporarily changes their size and adds header values, then ungroups the range and restores the saved widths before saving the workbook.
// Keywords: Aspose.Cells | C# | UngroupColumns | PreserveColumnWidth | GroupColumns | ColumnWidth | .NET Excel | Worksheet manipulation | Excel column grouping | code example
// Common Searches: Aspose.Cells ungroup columns keep width | restore column width after ungrouping C# | how to preserve Excel column sizes with Aspose.Cells | group then ungroup columns example Aspose.Cells .NET | C# code to save and reapply column widths in Excel
// Developer Intent: Remove column grouping while maintaining the original column dimensions.
// Use Cases: Temporarily widen columns for a header row, then revert to the template’s original layout. | Process a spreadsheet where columns are grouped for data entry, and the final report must match the original column widths. | Automate report generation that adjusts column grouping for visual emphasis without altering the final column sizing.
// AI Prompts: Write C# code using Aspose.Cells that ungroups columns A‑C and restores their previous widths after a temporary width change. | Explain how to capture column widths before grouping and reapply them after ungrouping with Aspose.Cells for .NET. | Create a reusable method that accepts a worksheet and a column range, ungroups the columns, and preserves their original widths.

using System;
using Aspose.Cells;

// Loads an Excel file, records the widths of a column range, groups the columns, temporarily changes their size and adds header values, then ungroups the range and restores the saved widths before saving the workbook.
class UngroupColumnsDemo
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        int firstCol = 0;   // first column index to ungroup
        int lastCol = 2;    // last column index to ungroup

        // Preserve original column widths so they stay consistent after ungrouping
        int colCount = lastCol - firstCol + 1;
        double[] originalWidths = new double[colCount];
        for (int i = 0; i < colCount; i++)
        {
            int colIndex = firstCol + i;
            originalWidths[i] = worksheet.Cells.Columns[colIndex].Width;
        }

        // Example modification: group the columns, change their width, add some data
        cells.GroupColumns(firstCol, lastCol);
        for (int i = firstCol; i <= lastCol; i++)
        {
            worksheet.Cells.Columns[i].Width = 30; // temporary width change
        }
        worksheet.Cells["A1"].PutValue("Grouped");
        worksheet.Cells["B1"].PutValue("Columns");
        worksheet.Cells["C1"].PutValue("Demo");

        // Ungroup the columns
        cells.UngroupColumns(firstCol, lastCol);

        // Restore the original widths to keep column width consistency
        for (int i = 0; i < colCount; i++)
        {
            int colIndex = firstCol + i;
            worksheet.Cells.Columns[colIndex].Width = originalWidths[i];
        }

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}
