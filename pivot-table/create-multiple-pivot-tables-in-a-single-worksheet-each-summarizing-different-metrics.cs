// Title: Create multiple pivot tables in a single worksheet with Aspose.Cells for .NET – sales by Category and quantity by SubCategory
// AI Prompts: Generate a pivot table at cell F3 that groups rows by the 'Category' field and sums the 'Sales' field, then apply the PivotTableStyleMedium9 style. | Create a second pivot table at cell F20 that groups rows by the 'SubCategory' field and sums the 'Quantity' field, apply the PivotTableStyleMedium4 style, refresh all pivot tables, and save the workbook as MultiplePivotTables.xlsx.
// Common Searches: asp.net add two pivot tables to the same sheet using Aspose.Cells C# | how to set different PivotTableStyleMedium styles for multiple pivot tables in Aspose.Cells | refresh all pivot tables programmatically after creation Aspose.Cells .NET | define source data range for pivot tables with formula reference Aspose.Cells | save workbook as xlsx after creating multiple pivot tables Aspose.Cells
// Tags: add multiple pivot tables Aspose.Cells | pivot table row field Category Aspose.Cells | pivot table data field Sales Aspose.Cells | pivot table row field SubCategory Aspose.Cells | pivot table data field Quantity Aspose.Cells | apply PivotTableStyleMedium9 Aspose.Cells | refresh pivot tables Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsMultiplePivotTables
{
    // The example creates a new workbook, populates it with Category, SubCategory, Sales, and Quantity data, defines the source range, adds a pivot table at F3 that sums Sales by Category with a medium style, adds another pivot table at F20 that sums Quantity by SubCategory with a different medium style, refreshes both pivot tables, and saves the file as MultiplePivotTables.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category, SubCategory, Sales, Quantity)
            Cells cells = sheet.Cells;
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("SubCategory");
            cells["C1"].PutValue("Sales");
            cells["D1"].PutValue("Quantity");

            // Sample rows
            string[] categories = { "Food", "Food", "Beverage", "Beverage", "Food", "Beverage" };
            string[] subCategories = { "Fruit", "Vegetable", "Soda", "Juice", "Fruit", "Juice" };
            double[] sales = { 1200, 800, 1500, 1100, 900, 1300 };
            int[] qty = { 30, 20, 50, 40, 25, 45 };

            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // Data starts from row 2
                cells[$"A{row}"].PutValue(categories[i]);
                cells[$"B{row}"].PutValue(subCategories[i]);
                cells[$"C{row}"].PutValue(sales[i]);
                cells[$"D{row}"].PutValue(qty[i]);
            }

            // Define the source data range for pivot tables
            // Using a formula style reference to the whole data block
            string sourceData = $"=Sheet1!{sheet.Cells.MaxDisplayRange.Address}";

            // Add first pivot table: Sales summary by Category
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex1 = pivots.Add(sourceData, "F3", "SalesByCategory");
            PivotTable pivot1 = pivots[pivotIndex1];
            pivot1.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot1.AddFieldToArea(PivotFieldType.Data, "Sales");
            // Optional: set a built‑in style
            pivot1.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Add second pivot table: Quantity summary by SubCategory
            int pivotIndex2 = pivots.Add(sourceData, "F20", "QtyBySubCategory");
            PivotTable pivot2 = pivots[pivotIndex2];
            pivot2.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivot2.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot2.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium4;

            // Refresh data for both pivot tables
            sheet.RefreshPivotTables();

            // Save the workbook
            workbook.Save("MultiplePivotTables.xlsx");
        }
    }
}
