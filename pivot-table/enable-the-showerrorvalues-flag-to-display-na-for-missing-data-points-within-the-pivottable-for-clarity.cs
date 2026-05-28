using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ShowErrorValuesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (including a missing value to illustrate the error display)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(null); // missing data point

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B3", "D5", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot fields
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

        // Enable custom error string display and set it to "#N/A"
        pivot.DisplayErrorString = true;
        pivot.ErrorString = "#N/A";

        // Refresh and calculate the pivot table to apply changes
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("ShowErrorValuesDemo.xlsx");
    }
}