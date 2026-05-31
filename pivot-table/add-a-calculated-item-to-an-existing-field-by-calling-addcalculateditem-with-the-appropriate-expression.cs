using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCalculatedItemDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["A5"].Value = "Vegetable";

            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table covering the data range and place it at D1
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Retrieve the row field (Category) to which we will add a calculated item
            PivotField categoryField = pivotTable.RowFields[0];

            // Add a calculated item that sums the amounts of Fruit and Vegetable
            // The formula uses the item names as they appear in the pivot field
            categoryField.AddCalculatedItem("Total_Fruit_Veg", "=Fruit + Vegetable");

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_With_CalculatedItem.xlsx");
        }
    }
}