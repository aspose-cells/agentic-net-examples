using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomListOrderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Item";
            sheet.Cells["C1"].Value = "Amount";

            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = "Apple";
            sheet.Cells["C2"].Value = 120;

            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["B3"].Value = "Banana";
            sheet.Cells["C3"].Value = 80;

            sheet.Cells["A4"].Value = "Fruit";
            sheet.Cells["B4"].Value = "Cherry";
            sheet.Cells["C4"].Value = 150;

            sheet.Cells["A5"].Value = "Vegetable";
            sheet.Cells["B5"].Value = "Carrot";
            sheet.Cells["C5"].Value = 60;

            sheet.Cells["A6"].Value = "Vegetable";
            sheet.Cells["B6"].Value = "Broccoli";
            sheet.Cells["C6"].Value = 90;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C6", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields: Category as row, Item as row (nested), Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Item");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable custom list sorting (optional, uses built‑in custom lists)
            pivotTable.CustomListSort = true;

            // Define a custom order for the "Item" field:
            // Desired order: Banana, Apple, Cherry, Broccoli, Carrot
            PivotField itemField = pivotTable.RowFields["Item"];
            PivotItemCollection items = itemField.PivotItems;

            // Set PositionInSameParentNode to control order within each parent (Category)
            // For "Fruit" category
            items["Banana"].PositionInSameParentNode = 0;
            items["Apple"].PositionInSameParentNode = 1;
            items["Cherry"].PositionInSameParentNode = 2;

            // For "Vegetable" category
            items["Broccoli"].PositionInSameParentNode = 0;
            items["Carrot"].PositionInSameParentNode = 1;

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomListOrder.xlsx");
        }
    }
}