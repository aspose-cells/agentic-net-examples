using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotFieldReorderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Category";
            cells["B1"].Value = "SubCategory";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Fruit";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 80;

            cells["A4"].Value = "Vegetable";
            cells["B4"].Value = "Carrot";
            cells["C4"].Value = 50;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = "Broccoli";
            cells["C5"].Value = 70;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add two fields to the row area: first "Category", then "SubCategory"
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");      // position 0
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory"); // position 1

            // Add the data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // At this point the hierarchy is Category -> SubCategory.
            // To change the display order to SubCategory -> Category, move the field.
            // Current positions: Category (0), SubCategory (1)
            // Move SubCategory (currPos = 1) to position 0.
            pivotTable.RowFields.Move(1, 0);

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldReorderResult.xlsx");
        }
    }
}