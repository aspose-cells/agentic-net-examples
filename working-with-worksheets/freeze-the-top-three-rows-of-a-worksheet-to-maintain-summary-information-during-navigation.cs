// Title: Freeze the first three rows of an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to freeze the first three rows in an Excel file with Aspose.Cells | Aspose.Cells FreezePanes method example for freezing header rows | How to keep top rows static while scrolling in Aspose.Cells .NET | Freeze top three rows of worksheet without freezing columns using Aspose.Cells | Set freeze panes for rows only in Aspose.Cells C#
// Tags: FreezePanes API Aspose.Cells C# | freeze rows only Aspose.Cells workbook | static header rows Excel Aspose.Cells | worksheet freeze panes without columns | Aspose.Cells top rows freeze example

using Aspose.Cells;
using System;

// The sample creates a new Workbook, adds sample data, and then freezes the top three rows (leaving columns unfrozen) by calling the four‑parameter FreezePanes method. The workbook is saved as FrozenTopRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Optional: add some sample data
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
                sheet.Cells[i, 1].PutValue(i * 10);
            }

            // Freeze the top three rows (no columns frozen)
            // Use the 4‑parameter overload required by recent Aspose.Cells versions
            int totalRows = sheet.Cells.MaxDataRow + 1;      // total rows in the pane
            int totalColumns = sheet.Cells.MaxDataColumn + 1; // total columns in the pane
            sheet.FreezePanes(3, 0, totalRows, totalColumns);

            // Save the workbook
            workbook.Save("FrozenTopRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
