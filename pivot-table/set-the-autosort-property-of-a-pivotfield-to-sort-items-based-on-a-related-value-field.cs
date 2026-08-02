using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotAutoSortDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Region";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "North";
            cells["B2"].Value = 1200;
            cells["A3"].Value = "South";
            cells["B3"].Value = 1500;
            cells["A4"].Value = "East";
            cells["B4"].Value = 800;
            cells["A5"].Value = "West";
            cells["B5"].Value = 1100;

            // Add a pivot table based on the data range
            int ptIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add the row field (Region) and the data field (Sales)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row field that will be auto‑sorted
            PivotField regionField = pivotTable.RowFields["Region"];

            // Enable auto‑sorting, set ascending order, and sort by the first data field (Sales)
            regionField.IsAutoSort = true;          // turn on auto sort
            regionField.IsAscendSort = true;        // sort ascending
            regionField.AutoSortField = 0;          // index of the data field to sort by (Sales)

            // Refresh the pivot table data and calculate results
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotFieldAutoSortDemo.xlsx");
        }
    }
}