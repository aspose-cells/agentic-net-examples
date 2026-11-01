// Title: Dynamically freeze top rows up to the last populated row with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Worksheet.Cells.MaxDataRow to locate the last non‑empty row and then calls FreezePanes to lock all rows above it. | Create a C# example that populates a worksheet, computes the maximum data row, and freezes both the header rows and the first column with Aspose.Cells. | Write a reusable C# method that accepts a Workbook and automatically applies FreezePanes based on the worksheet's MaxDataRow value.
// Common Searches: Aspose.Cells .NET freeze rows up to last data row programmatically | C# determine last populated row in Excel using MaxDataRow for freeze panes | How to set FreezePanes dynamically based on data range in Aspose.Cells | Freeze top header rows automatically with Aspose.Cells worksheet
// Tags: Aspose.Cells FreezePanes based on MaxDataRow | C# calculate last data row in Excel worksheet | dynamic row freezing in .NET Excel export | programmatic freeze panes for Excel using Aspose.Cells | Excel header freeze using Worksheet.Cells.MaxDataRow

using System;
using Aspose.Cells;

// The example creates a workbook, adds sample data, retrieves the index of the last row containing data via Cells.MaxDataRow, freezes all rows up to that index with FreezePanes, and saves the file as DynamicFreezeRows.xlsx.
class DynamicFreezeRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Sample data population (replace with actual data)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(30);
            // -------------------------------------------------

            // Determine the last row that contains data (zero‑based index)
            int maxDataRow = sheet.Cells.MaxDataRow;

            // Freeze rows up to the last data row (inclusive)
            // FreezePanes(firstVisibleRow, firstVisibleColumn, totalRowsToFreeze, totalColumnsToFreeze)
            sheet.FreezePanes(maxDataRow + 1, 0, maxDataRow + 1, 0);

            // Save the workbook
            workbook.Save("DynamicFreezeRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
