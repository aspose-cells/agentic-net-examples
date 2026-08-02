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
            Cells cells = sheet.Cells;

            // Populate sample data
            // Column A: Fruit, Column B: Quantity
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Quantity";

            cells["A2"].Value = "Apple";
            cells["B2"].Value = 10;

            cells["A3"].Value = "Banana";
            cells["B3"].Value = 20;

            cells["A4"].Value = "Orange";
            cells["B4"].Value = 15;

            cells["A5"].Value = "Pear";
            cells["B5"].Value = 5;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the Fruit field to the Row area and Quantity to the Data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Retrieve the row field (Fruit) and its PivotItems collection
            PivotField fruitField = pivotTable.RowFields["Fruit"];
            PivotItemCollection fruitItems = fruitField.PivotItems;

            // Define a custom order: Orange, Apple, Banana, Pear
            // Use PositionInSameParentNode to set the order within the same parent node
            // The first item gets position 0, the next 1, and so on.
            fruitItems["Orange"].PositionInSameParentNode = 0;
            fruitItems["Apple"].PositionInSameParentNode = 1;
            fruitItems["Banana"].PositionInSameParentNode = 2;
            fruitItems["Pear"].PositionInSameParentNode = 3;

            // Refresh and calculate the pivot table to apply the new order
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomListOrder.xlsx");
        }
    }
}