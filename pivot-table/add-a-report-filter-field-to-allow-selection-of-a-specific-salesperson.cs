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
            // Header row
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Salesperson";
            cells["C1"].Value = "Amount";

            // Data rows
            cells["A2"].Value = "Laptop";
            cells["B2"].Value = "Alice";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Laptop";
            cells["B3"].Value = "Bob";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Phone";
            cells["B4"].Value = "Alice";
            cells["C4"].Value = 800;

            cells["A5"].Value = "Phone";
            cells["B5"].Value = "Charlie";
            cells["C5"].Value = 950;

            // Define the source range for the pivot table
            string sourceRange = "A1:C5";

            // Add a pivot table at cell E3
            int pivotIndex = sheet.PivotTables.Add(sourceRange, "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields:
            // Row field: Product
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            // Data field: Amount (summed)
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            // Page (Report Filter) field: Salesperson
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Salesperson");

            // Optionally rename the page field for clarity
            pivotTable.PageFields[0].Name = "SalespersonFilter";

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Show the report filter page for the salesperson filter
            pivotTable.ShowReportFilterPageByName("SalespersonFilter");

            // Save the workbook
            workbook.Save("PivotTable_With_Salesperson_Filter.xlsx");
        }
    }
}