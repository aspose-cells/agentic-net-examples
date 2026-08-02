// Title: C# – Refresh All PivotTables Sequentially with Aspose.Cells for .NET
// Description: Load a workbook, iterate through each worksheet, call RefreshPivotTables() to update every pivot table, and save the file. This ensures all PivotTables are synchronized with the latest source data in a single pass.
// Keywords: Aspose.Cells refresh pivot tables | C# refresh all pivot tables | RefreshPivotTables method | sequential pivot table update | Excel workbook pivot refresh .NET | global Excel automation
// Common Searches: how to refresh every pivot table in an Excel file using Aspose.Cells | C# code to refresh pivot tables worksheet by worksheet | update all pivot tables after data change Aspose.Cells | programmatically refresh pivot tables in a workbook
// Developer Intent: Update every PivotTable in a workbook so it reflects the current data source.
// Use Cases: Refresh all PivotTables after bulk data import before generating a financial report. | Automate pivot table synchronization in multi‑sheet dashboards for nightly builds. | Ensure consistency of PivotTables across worksheets in a template used by multiple users.
// AI Prompts: Write C# code with Aspose.Cells that refreshes pivot tables only on selected worksheets. | Show how to catch exceptions from RefreshPivotTables, log the worksheet name, and continue processing. | Provide an example that refreshes all pivot tables, recalculates formulas, and then saves the workbook.

using System;
using Aspose.Cells;

// Load a workbook, iterate through each worksheet, call RefreshPivotTables() to update every pivot table, and save the file. This ensures all PivotTables are synchronized with the latest source data in a single pass.
class RefreshAllPivotTables
{
    static void Main()
    {
        // Load the workbook that contains the pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Refresh pivot tables worksheet by worksheet (sequentially)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Refresh all pivot tables in the current worksheet
            sheet.RefreshPivotTables();
        }

        // Save the workbook after all pivot tables have been refreshed
        workbook.Save("output.xlsx");
    }
}
