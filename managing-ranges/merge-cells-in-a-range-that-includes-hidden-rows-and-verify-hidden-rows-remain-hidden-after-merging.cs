// Title: Merge a cell range that includes hidden rows and ensure the rows stay hidden with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge range that contains hidden rows without unhiding them | how to keep rows hidden after merging cells using Aspose.Cells .NET | Aspose.Cells merge cells A1:C3 while preserving hidden row properties | verify hidden rows remain hidden after cell merge in Aspose.Cells C# | merge cells across hidden rows programmatically with Aspose.Cells
// Tags: merge cells hidden rows Aspose.Cells | preserve hidden row property Aspose.Cells | Aspose.Cells merge range with hidden rows | C# check hidden rows post merge | save workbook after merging hidden rows Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, hides specific rows, merges a range that spans those hidden rows, then confirms the rows stay hidden after the merge and saves the file as MergedHiddenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data
            sheet.Cells["A1"].PutValue("R1C1");
            sheet.Cells["B1"].PutValue("R1C2");
            sheet.Cells["C1"].PutValue("R1C3");
            sheet.Cells["A2"].PutValue("R2C1");
            sheet.Cells["B2"].PutValue("R2C2");
            sheet.Cells["C2"].PutValue("R2C3");
            sheet.Cells["A3"].PutValue("R3C1");
            sheet.Cells["B3"].PutValue("R3C2");
            sheet.Cells["C3"].PutValue("R3C3");

            // Hide rows 2 and 3 (zero‑based indices 1 and 2)
            sheet.Cells.Rows[1].IsHidden = true;
            sheet.Cells.Rows[2].IsHidden = true;

            // Merge cells A1:C3 which includes the hidden rows
            // Parameters: startRow, startColumn, totalRows, totalColumns
            sheet.Cells.Merge(0, 0, 3, 3);

            // Verify that the hidden rows are still hidden after merging
            bool isRow2Hidden = sheet.Cells.Rows[1].IsHidden;
            bool isRow3Hidden = sheet.Cells.Rows[2].IsHidden;

            Console.WriteLine($"Row 2 hidden after merge: {isRow2Hidden}");
            Console.WriteLine($"Row 3 hidden after merge: {isRow3Hidden}");

            // Save the workbook (optional, for visual verification)
            workbook.Save("MergedHiddenRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
