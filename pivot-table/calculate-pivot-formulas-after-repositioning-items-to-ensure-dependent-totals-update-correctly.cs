// Title: Update Pivot Table Totals After Relocating It with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to move a pivot table from one cell range to another using Aspose.Cells for .NET, then invoke CalculateRange, RefreshData, and CalculateData to recalculate totals and keep dependent formulas accurate before saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | pivot table move | PivotTable.MoveTo | CalculateRange | RefreshData | CalculateData | recalculate pivot | update pivot totals | reposition pivot table | Excel automation | workbook save
// Common Searches: how to refresh a pivot table after moving it with Aspose.Cells | Aspose.Cells recalculate totals after pivot relocation | C# move pivot table and update formulas | refresh all pivot tables in a workbook Aspose | update dependent totals when pivot address changes
// Developer Intent: Programmatically move a pivot table to a new address and recalculate its data so that all subtotals and grand totals remain correct.
// Use Cases: Shift a sales pivot from C3 to E5 in a financial report and ensure totals reflect the new position. | Refresh multiple pivots after a dynamic layout change in a dashboard workbook. | Automate pivot repositioning based on user input while preserving accurate calculations.
// AI Prompts: Generate C# code that moves an Aspose.Cells pivot table to a specified cell and then refreshes its calculations. | Explain why CalculateRange, RefreshData, and CalculateData are required after PivotTable.MoveTo. | Create a reusable method that accepts a pivot name and target address, moves the pivot, and updates all dependent totals.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to move a pivot table from one cell range to another using Aspose.Cells for .NET, then invoke CalculateRange, RefreshData, and CalculateData to recalculate totals and keep dependent formulas accurate before saving the workbook.
class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new workbook and add sample data
        // -------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Header
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Amount");

        // Sample rows
        string[] categories = { "A", "B", "A", "C", "B", "C", "A" };
        for (int i = 0; i < categories.Length; i++)
        {
            dataSheet.Cells[i + 1, 0].PutValue(categories[i]);          // Column A
            dataSheet.Cells[i + 1, 1].PutValue((i + 1) * 100);        // Column B
        }

        // -------------------------------------------------
        // 2. Create a pivot table on a separate sheet
        // -------------------------------------------------
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
        // Define source range (including header row)
        string sourceRange = $"A1:B{categories.Length + 1}";
        // Place pivot table initially at C3
        int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "C3", "SalesPivot");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Add fields: Category as row, Amount as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // -------------------------------------------------
        // 3. Initial calculation of the pivot table
        // -------------------------------------------------
        pivotTable.CalculateRange();   // Ensure the pivot range is correct
        pivotTable.RefreshData();      // Pull data from the source
        pivotTable.CalculateData();    // Compute totals, subtotals, etc.

        // -------------------------------------------------
        // 4. Reposition the pivot table to a new location
        // -------------------------------------------------
        // Move the whole pivot table to cell E5
        pivotTable.MoveTo("E5");

        // -------------------------------------------------
        // 5. Re‑calculate after moving to update dependent totals
        // -------------------------------------------------
        pivotTable.CalculateRange();   // Re‑evaluate the pivot's range after move
        pivotTable.RefreshData();      // Refresh data source (necessary after move)
        pivotTable.CalculateData();    // Re‑calculate totals so dependent formulas are correct

        // Optional: refresh all pivot tables in the workbook (covers any other pivots)
        workbook.Worksheets.RefreshPivotTables();

        // -------------------------------------------------
        // 6. Save the workbook
        // -------------------------------------------------
        workbook.Save("PivotRepositioned.xlsx");
    }
}
