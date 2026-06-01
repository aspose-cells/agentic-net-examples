using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: Product, Region, UnitsSold
        cells["A1"].PutValue("Product");
        cells["B1"].PutValue("Region");
        cells["C1"].PutValue("UnitsSold");

        cells["A2"].PutValue("Apple");  cells["B2"].PutValue("North"); cells["C2"].PutValue(120);
        cells["A3"].PutValue("Apple");  cells["B3"].PutValue("South"); cells["C3"].PutValue(80);
        cells["A4"].PutValue("Banana"); cells["B4"].PutValue("North"); cells["C4"].PutValue(150);
        cells["A5"].PutValue("Banana"); cells["B5"].PutValue("South"); cells["C5"].PutValue(130);

        // Add a pivot table covering the data range and place it at E3
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add row fields to group by Product and Region
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");

        // Add the UnitsSold field as a data field and enforce Sum aggregation
        int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "UnitsSold");
        PivotField unitsSoldField = pivotTable.DataFields[dataFieldPos];
        unitsSoldField.Function = ConsolidationFunction.Sum; // Sum aggregation

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the pivot table
        workbook.Save("TotalUnitsSoldPivot.xlsx");
    }
}