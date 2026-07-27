using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsReportFilterDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with a Salesperson column
            cells["A1"].Value = "Salesperson";
            cells["B1"].Value = "Product";
            cells["C1"].Value = "Quantity";

            cells["A2"].Value = "Alice";
            cells["B2"].Value = "Apple";
            cells["C2"].Value = 10;

            cells["A3"].Value = "Bob";
            cells["B3"].Value = "Banana";
            cells["C3"].Value = 15;

            cells["A4"].Value = "Alice";
            cells["B4"].Value = "Orange";
            cells["C4"].Value = 20;

            cells["A5"].Value = "Bob";
            cells["B5"].Value = "Apple";
            cells["C5"].Value = 5;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields: Salesperson as a report filter (page field)
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Salesperson");
            // Optionally rename the page field for clarity
            pivotTable.PageFields[0].Name = "SalespersonFilter";

            // Add row field (Product) and data field (Quantity)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Show the report filter page for the salesperson filter
            pivotTable.ShowReportFilterPageByName("SalespersonFilter");

            // Save the workbook
            workbook.Save("ReportFilterSalespersonDemo.xlsx");
        }
    }
}