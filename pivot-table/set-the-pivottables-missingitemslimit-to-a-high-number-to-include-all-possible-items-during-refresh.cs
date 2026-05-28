using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Year";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Fruit";
            cells["B2"].Value = 2020;
            cells["C2"].Value = 50;

            cells["A3"].Value = "Vegetable";
            cells["B3"].Value = 2020;
            cells["C3"].Value = 60;

            cells["A4"].Value = "Fruit";
            cells["B4"].Value = 2021;
            cells["C4"].Value = 70;

            cells["A5"].Value = "Vegetable";
            cells["B5"].Value = 2021;
            cells["C5"].Value = 80;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("=Sheet1!A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Category
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Year
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Amount

            // Set MissingItemsLimit to Max to retain all possible items during refresh
            pivotTable.MissingItemsLimit = PivotMissingItemLimitType.Max;

            // Refresh the pivot cache and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_MissingItemsLimit_Max.xlsx");
        }
    }
}