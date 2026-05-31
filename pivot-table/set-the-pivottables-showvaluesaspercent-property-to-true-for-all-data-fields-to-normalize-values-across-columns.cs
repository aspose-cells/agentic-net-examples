using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
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

            cells["A2"].Value = "Food";
            cells["B2"].Value = "Fruit";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Food";
            cells["B3"].Value = "Vegetable";
            cells["C3"].Value = 800;

            cells["A4"].Value = "Beverage";
            cells["B4"].Value = "Tea";
            cells["C4"].Value = 500;

            cells["A5"].Value = "Beverage";
            cells["B5"].Value = "Coffee";
            cells["C5"].Value = 700;

            // Add a pivot table to the worksheet
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivots[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Set ShowValuesAsPercent (percentage of total) for all data fields
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                // Use ShowValuesSetting.CalculationType to display values as percentage of total
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.PercentageOfTotal;
            }

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_ShowValuesAsPercent.xlsx");
        }
    }
}