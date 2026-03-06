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

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(3000);

        // Add a pivot table covering the data range and place it at E3
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add the product column as a row field
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");

        // Add the sales column as a data field
        int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        PivotField dataField = pivot.DataFields[dataFieldPos];
        dataField.Function = ConsolidationFunction.Sum;

        // Set a custom number format for the data field (e.g., currency)
        dataField.NumberFormat = "$#,##0.00";

        // Refresh and calculate the pivot table to apply the changes
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("PivotNumberFormatDemo.xlsx", SaveFormat.Xlsx);
    }
}