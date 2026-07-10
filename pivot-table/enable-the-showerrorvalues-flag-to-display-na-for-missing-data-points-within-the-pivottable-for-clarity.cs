using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ShowErrorValuesPivotDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data; include a missing value for demonstration
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        // B4 is left blank to represent a missing data point

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot fields: Category as row, Value as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Value

        // Enable the display of error strings and set the custom error text to "#N/A"
        pivotTable.DisplayErrorString = true;
        pivotTable.ErrorString = "#N/A";

        // Refresh the source data and calculate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the configured pivot table
        workbook.Save("PivotTableShowErrorValues.xlsx");
    }
}