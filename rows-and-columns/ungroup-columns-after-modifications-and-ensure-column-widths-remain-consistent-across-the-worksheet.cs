// Title: How to ungroup columns while preserving their original widths using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that groups a range of columns, modifies cell values, then ungroups the columns and restores each column's original width with Aspose.Cells. | Demonstrate storing column widths before grouping and reapplying them after ungrouping in an Aspose.Cells workbook.
// Common Searches: asp.net ungroup columns keep original width Aspose.Cells | c# Aspose.Cells restore column widths after ungrouping columns | how to maintain column width consistency when ungrouping columns in an Excel workbook using Aspose.Cells | preserve column width after grouping and ungrouping columns with Aspose.Cells .NET | example of ungrouping columns without changing column width in Aspose.Cells
// Tags: ungroup columns Aspose.Cells C# | preserve column width Aspose.Cells | column width consistency after ungrouping | group columns Aspose.Cells example | restore original column widths .NET Excel

using System;
using Aspose.Cells;

// The example creates a workbook, sets specific column widths, groups columns 0‑2, updates cell data while grouped, then ungroups the columns and reapplies the stored widths to keep the layout unchanged before saving as UngroupColumnsConsistentWidth.xlsx.
class UngroupColumnsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data in columns A, B and C
        cells["A1"].PutValue("Short");
        cells["B1"].PutValue("Medium length text");
        cells["C1"].PutValue("Very very long text that needs column width adjustment");

        // Set initial column widths for demonstration
        cells.Columns[0].Width = 12; // Column A
        cells.Columns[1].Width = 20; // Column B
        cells.Columns[2].Width = 30; // Column C

        // Store the original widths so they can be restored after ungrouping
        double[] originalWidths = new double[3];
        for (int i = 0; i < 3; i++)
        {
            originalWidths[i] = cells.Columns[i].Width;
        }

        // Group columns 0‑2 and hide them (optional, just to show that ungroup works)
        cells.GroupColumns(0, 2, true);

        // Perform some modifications while the columns are grouped
        cells["A2"].PutValue("Additional data");
        cells["B2"].PutValue("More data");
        cells["C2"].PutValue("Even more data");

        // Ungroup the columns (0 to 2)
        cells.UngroupColumns(0, 2);

        // Restore the original column widths to keep them consistent across the worksheet
        for (int i = 0; i < 3; i++)
        {
            cells.Columns[i].Width = originalWidths[i];
        }

        // Save the workbook
        workbook.Save("UngroupColumnsConsistentWidth.xlsx");
    }
}
