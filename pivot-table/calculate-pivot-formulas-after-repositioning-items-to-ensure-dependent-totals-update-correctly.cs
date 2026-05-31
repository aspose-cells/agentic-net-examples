using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRecalc
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source data for the pivot table
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("SubCategory");
            dataSheet.Cells["C1"].PutValue("Amount");

            // Sample rows
            string[] categories = { "Fruit", "Fruit", "Vegetable", "Vegetable", "Fruit", "Vegetable" };
            string[] subCategories = { "Apple", "Banana", "Carrot", "Broccoli", "Orange", "Spinach" };
            double[] amounts = { 120, 80, 150, 200, 90, 110 };

            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // data starts at row 2
                dataSheet.Cells[$"A{row}"].PutValue(categories[i]);
                dataSheet.Cells[$"B{row}"].PutValue(subCategories[i]);
                dataSheet.Cells[$"C{row}"].PutValue(amounts[i]);
            }

            // -------------------------------------------------
            // 2. Add a pivot table based on the source data
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            // Define source range (including header)
            string sourceRange = "Data!A1:C7";
            // Destination cell for the pivot table
            string destCell = "A3";

            int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, destCell, "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category (row), SubCategory (column), Amount (data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Initial calculation so the pivot has data
            pivotTable.CalculateRange();
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 3. Reposition pivot items (move the whole pivot table)
            // -------------------------------------------------
            // Move the pivot table to a new location within the same worksheet
            pivotTable.MoveTo("E10"); // alternative: pivotTable.MoveTo(9, 4);

            // -------------------------------------------------
            // 4. Re‑calculate dependent totals after repositioning
            // -------------------------------------------------
            // Re‑calculate the pivot range (in case the move changed the occupied area)
            pivotTable.CalculateRange();

            // Refresh the underlying cache from the source data (necessary after any structural change)
            pivotTable.RefreshData();

            // Re‑calculate the pivot data so totals reflect the new layout
            pivotTable.CalculateData();

            // Optionally refresh all pivot tables in the workbook (covers scenarios with multiple pivots)
            workbook.Worksheets.RefreshAll();

            // -------------------------------------------------
            // 5. Save the workbook
            // -------------------------------------------------
            workbook.Save("PivotRecalcAfterMove.xlsx");
        }
    }
}