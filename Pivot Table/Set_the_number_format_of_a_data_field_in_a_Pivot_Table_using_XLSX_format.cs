using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotNumberFormatDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
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

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add a row field (Product) and a data field (Sales)
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the data field and set a custom number format
            PivotField dataField = pivot.DataFields[dataFieldPos];
            dataField.Function = ConsolidationFunction.Sum; // optional: define aggregation
            dataField.NumberFormat = "$#,##0.00"; // custom currency format

            // Refresh and calculate the pivot table to apply changes
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotFieldNumberFormatDemo.xlsx");
        }
    }
}