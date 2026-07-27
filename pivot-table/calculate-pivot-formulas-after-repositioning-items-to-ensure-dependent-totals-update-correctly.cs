using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample source data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("Vegetable");
            dataSheet.Cells["B3"].PutValue(800);
            dataSheet.Cells["A4"].PutValue("Fruit");
            dataSheet.Cells["B4"].PutValue(1500);
            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue(900);

            // Add a second worksheet that will host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table (source range, destination cell, name)
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "C3", "SalesPivot");
            PivotTable pivot = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category to rows, Amount to data
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Initial calculation of the pivot table range and data
            pivot.CalculateRange();   // Ensure the range reflects the source data
            pivot.RefreshData();      // Pull latest source data into the pivot cache
            pivot.CalculateData();    // Populate the pivot table cells

            // Move the pivot table to a new location (repositioning items)
            pivot.MoveTo("E10");      // New top‑left cell for the pivot table

            // After moving, recalculate range and data so dependent totals update correctly
            pivot.CalculateRange();   // Re‑evaluate the pivot's range after relocation
            pivot.RefreshData();      // Refresh cache in case source data changed
            pivot.CalculateData();    // Re‑calculate totals and formulas

            // Optionally refresh all pivot tables in the workbook (covers any other pivots)
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook (lifecycle: save)
            workbook.Save("PivotRefreshAfterMove.xlsx");
        }
    }
}