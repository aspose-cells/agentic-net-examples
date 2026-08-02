// Title: Refresh Pivot Tables Programmatically with Aspose.Cells for .NET (C#)
// Description: Load a workbook, modify the source range of a pivot table, invoke Worksheet.RefreshPivotTables() to recalculate all pivots on the sheet, and save the updated file. This example demonstrates how to keep pivot reports in sync after data changes using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | RefreshPivotTables | pivot table refresh | update pivot after data change | programmatic pivot update | worksheet.RefreshPivotTables method | Excel automation
// Common Searches: Aspose.Cells refresh pivot tables C# | how to update pivot after editing cells .NET | Worksheet.RefreshPivotTables example | programmatically recalculate Excel pivots | refresh all pivots in a worksheet using Aspose
// Developer Intent: Recalculate pivot tables so they reflect modified source data before saving the workbook.
// Use Cases: After adjusting sales figures in the source sheet, call RefreshPivotTables to produce an up‑to‑date summary report. | When importing external CSV data into a workbook, use the method to ensure any existing pivots display the new totals. | In an automated nightly job that alters multiple worksheets, invoke RefreshPivotTables on each sheet to keep all analyses current.
// AI Prompts: Show how to refresh a single pivot table by name with Aspose.Cells. | Demonstrate refreshing pivots after appending new rows to the source range in C#. | Explain error handling for RefreshPivotTables when the data source is missing or corrupted.

using System;
using Aspose.Cells;

// Load a workbook, modify the source range of a pivot table, invoke Worksheet.RefreshPivotTables() to recalculate all pivots on the sheet, and save the updated file. This example demonstrates how to keep pivot reports in sync after data changes using Aspose.Cells.
class RefreshPivotExample
{
    static void Main()
    {
        // Load an existing workbook that contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the first worksheet holds the source data and the pivot table
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify the source data that the pivot table is based on
        worksheet.Cells["B2"].PutValue(1500);
        worksheet.Cells["B3"].PutValue(2500);

        // Refresh all pivot tables in this worksheet so they reflect the new data
        worksheet.RefreshPivotTables();

        // Save the workbook with the refreshed pivot tables
        workbook.Save("output.xlsx");
    }
}
