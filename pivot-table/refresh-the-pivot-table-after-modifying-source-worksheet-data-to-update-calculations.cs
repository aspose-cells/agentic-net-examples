// Title: Refresh Pivot Tables After Changing Source Data with Aspose.Cells for .NET
// Description: Loads a workbook, updates cells in the source worksheet, calls Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table, and saves the result. Demonstrates how to keep pivot‑table summaries in sync with modified data using Aspose.Cells for C#.
// Keywords: Aspose.Cells refresh pivot tables | C# RefreshPivotTables example | update pivot cache .NET | pivot table recalculation Aspose | modify source data and refresh pivot | Aspose.Cells workbook pivot refresh
// Common Searches: how to refresh all pivot tables in Aspose.Cells | RefreshPivotTables method C# | update pivot tables after editing source cells Aspose | programmatic pivot table refresh .NET | Aspose.Cells pivot table recalc after data change
// Developer Intent: Programmatically refresh every pivot table in a workbook so it reflects recent changes to the source data.
// Use Cases: Adjust daily sales figures in the data sheet, refresh all pivot tables, and generate an up‑to‑date sales report. | Automate a nightly inventory update that modifies stock levels, triggers a pivot refresh, and saves the refreshed workbook for downstream systems. | Change financial metrics in a budgeting worksheet and ensure all pivot summaries across multiple sheets display the new values before distribution.
// AI Prompts: Show how to refresh a single pivot table by name after editing source cells using Aspose.Cells for .NET. | Provide robust error handling for RefreshPivotTables when processing large workbooks with many pivots. | Explain how to refresh pivot tables that span multiple worksheets, each with its own data source, using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads a workbook, updates cells in the source worksheet, calls Workbook.Worksheets.RefreshPivotTables() to recalculate every pivot table, and saves the result. Demonstrates how to keep pivot‑table summaries in sync with modified data using Aspose.Cells for C#.
class RefreshPivotDemo
{
    static void Main()
    {
        // Load an existing workbook that contains pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that holds the source data for the pivot tables
        Worksheet dataSheet = workbook.Worksheets[0];

        // Modify some source data values
        dataSheet.Cells["B2"].PutValue(1500);
        dataSheet.Cells["B3"].PutValue(2500);

        // Refresh all pivot tables in the entire workbook to reflect the changes
        workbook.Worksheets.RefreshPivotTables();

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
