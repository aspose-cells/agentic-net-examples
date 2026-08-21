// Title: Refresh All PivotTables in an Excel Workbook Using Aspose.Cells for .NET (C#)
// Description: Load a workbook, call WorksheetCollection.RefreshPivotTables() to update every PivotTable on every sheet, and save the file. This ensures that all pivot caches are synchronized before the workbook is exported or further processed.
// Keywords: Aspose.Cells refresh pivot tables | WorksheetCollection.RefreshPivotTables | C# update all pivot tables | synchronize pivot cache Excel | refresh pivot tables programmatically | .NET Excel pivot refresh
// Common Searches: how to refresh all pivot tables with Aspose.Cells | Aspose.Cells C# refresh pivot tables across worksheets | refresh pivot tables before saving workbook .NET | WorksheetCollection.RefreshPivotTables example
// Developer Intent: Update every PivotTable in a workbook so the data reflects the latest source before saving or exporting.
// Use Cases: Refresh all pivots after modifying source data to keep multiple reports consistent. | Automate nightly jobs that regenerate Excel files with up‑to‑date pivot calculations. | Prepare a workbook for PDF conversion, ensuring pivot tables display current values.
// AI Prompts: Show how to refresh selected PivotTables while leaving others unchanged using Aspose.Cells. | Give an example that refreshes PivotTables and then recalculates all formulas in the same workbook. | Explain error handling for RefreshPivotTables when the data source is missing or corrupted.

using System;
using Aspose.Cells;

// Load a workbook, call WorksheetCollection.RefreshPivotTables() to update every PivotTable on every sheet, and save the file. This ensures that all pivot caches are synchronized before the workbook is exported or further processed.
class Program
{
    static void Main()
    {
        // Load the existing workbook that contains one or more PivotTables
        Workbook workbook = new Workbook("input.xlsx");

        // Refresh all PivotTables across all worksheets in the workbook.
        // This uses the WorksheetCollection.RefreshPivotTables() method.
        workbook.Worksheets.RefreshPivotTables();

        // Save the workbook after the refresh operation.
        workbook.Save("output.xlsx");
    }
}
