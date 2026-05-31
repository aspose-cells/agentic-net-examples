using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsMultiplePivotTables
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet (will hold both data and pivot tables)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:D11)
            // Columns: Category, SubCategory, Sales, Quantity
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["C1"].PutValue("Sales");
            sheet.Cells["D1"].PutValue("Quantity");

            string[] categories = { "Food", "Food", "Food", "Beverage", "Beverage", "Beverage" };
            string[] subCategories = { "Fruit", "Vegetable", "Grain", "Soda", "Juice", "Water" };
            double[] sales = { 1200, 800, 500, 1500, 1100, 900 };
            int[] qty = { 30, 20, 15, 40, 35, 25 };

            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // data starts at row 2
                sheet.Cells[row, 0].PutValue(categories[i]);
                sheet.Cells[row, 1].PutValue(subCategories[i]);
                sheet.Cells[row, 2].PutValue(sales[i]);
                sheet.Cells[row, 3].PutValue(qty[i]);
            }

            // Define the source data range for the pivot tables
            // Using a formula style reference to the whole data block
            string sourceData = $"=Sheet1!{sheet.Cells.MaxDisplayRange.Address}";

            // -------------------------------------------------
            // First PivotTable: Total Sales by Category
            // -------------------------------------------------
            // Destination cell for the first pivot table (placed below the data)
            string destCell1 = "F2";
            string tableName1 = "SalesByCategory";

            int pivotIndex1 = sheet.PivotTables.Add(sourceData, destCell1, tableName1);
            PivotTable pivot1 = sheet.PivotTables[pivotIndex1];

            // Configure fields: Category as Row, Sales as Data
            pivot1.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot1.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Optional: apply a built‑in style
            pivot1.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // -------------------------------------------------
            // Second PivotTable: Total Quantity by SubCategory
            // -------------------------------------------------
            // Destination cell for the second pivot table (placed further down)
            string destCell2 = "F20";
            string tableName2 = "QtyBySubCategory";

            int pivotIndex2 = sheet.PivotTables.Add(sourceData, destCell2, tableName2);
            PivotTable pivot2 = sheet.PivotTables[pivotIndex2];

            // Configure fields: SubCategory as Row, Quantity as Data
            pivot2.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivot2.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Apply a different style for visual distinction
            pivot2.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium12;

            // Refresh and calculate data for both pivot tables
            pivot1.RefreshData();
            pivot1.CalculateData();

            pivot2.RefreshData();
            pivot2.CalculateData();

            // Save the workbook with the created pivot tables
            workbook.Save("MultiplePivotTables.xlsx");
        }
    }
}