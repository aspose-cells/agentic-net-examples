using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsHierarchicalGroupingDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            // Header row
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["C1"].PutValue("Sales");
            sheet.Cells["D1"].PutValue("Quantity");

            // Sample rows
            DateTime[] dates = {
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 15),
                new DateTime(2023, 2, 10),
                new DateTime(2023, 2, 20),
                new DateTime(2023, 3, 5),
                new DateTime(2023, 3, 25)
            };
            string[] categories = { "Electronics", "Furniture", "Electronics", "Furniture", "Electronics", "Furniture" };
            double[] sales = { 1200, 800, 1500, 950, 1100, 700 };
            int[] qty = { 3, 2, 5, 4, 2, 1 };

            for (int i = 0; i < dates.Length; i++)
            {
                int row = i + 2; // data starts at row 2 (1‑based)
                sheet.Cells[row, 0].PutValue(dates[i]);
                sheet.Cells[row, 1].PutValue(categories[i]);
                sheet.Cells[row, 2].PutValue(sales[i]);
                sheet.Cells[row, 3].PutValue(qty[i]);
            }

            // ---------- Create a pivot table ----------
            // Data range: A1:D7 (including header)
            int pivotIndex = sheet.PivotTables.Add("A1:D7", "F3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add Date field to the Row area (will be grouped)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add Category field to the Column area (second column aggregation)
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Category");

            // Add Sales as Data field (sum)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // ---------- Hierarchical grouping on the Date field ----------
            // Retrieve the Date pivot field (first row field)
            PivotField dateField = pivotTable.RowFields[0];

            // Define grouping period (full year 2023) and multiple group types:
            // First group by Years, then by Months (nested hierarchy)
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);
            PivotGroupByType[] groupTypes = new PivotGroupByType[]
            {
                PivotGroupByType.Years,   // outer level
                PivotGroupByType.Months   // inner level
            };

            // Apply hierarchical grouping (interval = 1, do not create a new field for the first group)
            dateField.GroupBy(startDate, endDate, groupTypes, 1, false);

            // ---------- Optional: Group numeric field (Quantity) ----------
            // Add Quantity as another Data field to demonstrate two‑column aggregation
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Retrieve the Quantity pivot field (first base field with numeric values)
            // Note: BaseFields[2] corresponds to the "Quantity" column (0‑based index)
            PivotField qtyField = pivotTable.BaseFields[3];
            // Group Quantity by intervals of 2 (e.g., 0‑2, 2‑4, etc.) without creating a new field
            qtyField.GroupBy(2.0, false);

            // ---------- Refresh and calculate the pivot table ----------
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ---------- Save the workbook ----------
            workbook.Save("HierarchicalGroupingDemo.xlsx");
        }
    }
}