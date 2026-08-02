// Title: Aspose.Cells for .NET – Auto‑Refresh PivotTable Summary Sheet (Macro‑Style)
// Description: C# example that builds a workbook with a data worksheet, adds a PivotTable on a separate Summary sheet, sets ManualUpdate = false for automatic refresh, modifies source cells, calls RefreshPivotTables and CalculateFormula, and saves both the initial and updated files.
// Keywords: Aspose.Cells | .NET PivotTable | auto refresh pivot | ManualUpdate false | RefreshPivotTables | summary sheet | Excel macro replacement | programmatic pivot refresh | C# workbook example | dynamic summary table
// Common Searches: Aspose.Cells auto refresh pivot table .NET | How to update PivotTable after data change using Aspose.Cells | Set ManualUpdate property in Aspose.Cells PivotTable | Refresh all pivot tables in a workbook C# | Create summary sheet with PivotTable programmatically
// Developer Intent: Create a PivotTable that updates automatically when its source data is edited.
// Use Cases: Financial reporting workbook where expense totals adjust instantly as rows are edited. | Inventory dashboard that reflects real‑time stock levels after batch data imports. | Excel macro‑free solution that keeps charts and formulas synchronized after programmatic data changes.
// AI Prompts: Generate C# code with Aspose.Cells to build a PivotTable that auto‑refreshes when source data changes, including ManualUpdate configuration and RefreshPivotTables call. | Show how to modify cells in a worksheet and then refresh the associated PivotTable and recalculate formulas using Aspose.Cells for .NET. | Explain how to mimic an Excel macro that updates a summary sheet by configuring a PivotTable for automatic refresh in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that builds a workbook with a data worksheet, adds a PivotTable on a separate Summary sheet, sets ManualUpdate = false for automatic refresh, modifies source cells, calls RefreshPivotTables and CalculateFormula, and saves both the initial and updated files.
class SummaryUpdater
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // -----------------------------------------------------------------
        // 1. Prepare source data sheet
        // -----------------------------------------------------------------
        Worksheet dataSheet = wb.Worksheets[0];
        dataSheet.Name = "Data";

        // Header
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");

        // Sample rows
        dataSheet.Cells["A2"].PutValue("Food");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Transport");
        dataSheet.Cells["B3"].PutValue(50);
        dataSheet.Cells["A4"].PutValue("Food");
        dataSheet.Cells["B4"].PutValue(150);
        dataSheet.Cells["A5"].PutValue("Utilities");
        dataSheet.Cells["B5"].PutValue(80);

        // -----------------------------------------------------------------
        // 2. Create a summary sheet with a PivotTable (acts as the macro)
        // -----------------------------------------------------------------
        Worksheet summarySheet = wb.Worksheets.Add("Summary");

        // Define the source range for the pivot table
        string sourceRange = "Data!A1:B5";

        // Add the pivot table to the summary sheet, top‑left cell is A1
        int pivotIdx = summarySheet.PivotTables.Add(sourceRange, "A1", "SummaryPivot");
        PivotTable pivot = summarySheet.PivotTables[pivotIdx];

        // Configure the pivot: rows = Category, data = Sum of Amount
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Ensure the pivot updates automatically when source data changes
        // (ManualUpdate = false is the default, set explicitly for clarity)
        pivot.ManualUpdate = false;

        // Save the workbook with the initial summary
        wb.Save("SummaryWorkbook.xlsx");

        // -----------------------------------------------------------------
        // 3. Simulate a change in the source data
        // -----------------------------------------------------------------
        dataSheet.Cells["B3"].PutValue(70); // Updated Transport amount

        // Refresh all pivot tables so the summary reflects the new data
        wb.Worksheets.RefreshPivotTables();

        // Recalculate any other formulas that might exist
        wb.CalculateFormula();

        // Save the workbook after the automatic update
        wb.Save("SummaryWorkbook_Updated.xlsx");
    }
}
