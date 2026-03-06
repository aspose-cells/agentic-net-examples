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
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add a row field (Product)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add a data field (Sales) and set its number format
            int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            PivotField dataField = pivotTable.DataFields[dataFieldIndex];
            dataField.Function = ConsolidationFunction.Sum;          // Sum the sales values
            dataField.NumberFormat = "$#,##0.00";                    // Custom number format for currency

            // Refresh and calculate the pivot table to apply the changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to an XLSX file
            workbook.Save("PivotFieldNumberFormatDemo.xlsx");
        }
    }
}