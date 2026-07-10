using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Columns: Product | Sales
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Sales";

            sheet.Cells["A2"].Value = "B";
            sheet.Cells["A3"].Value = "C";
            sheet.Cells["A4"].Value = "A";

            sheet.Cells["B2"].Value = 150;
            sheet.Cells["B3"].Value = 300;
            sheet.Cells["B4"].Value = 200;

            // Add a pivot table based on the data range A1:B4, place it at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Product" field as a row field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the "Sales" field as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Retrieve the row field that we just added
            PivotField productField = pivotTable.RowFields[0];

            // Enable automatic sorting for the field
            productField.IsAutoSort = true;

            // Set descending order (IsAscendSort = false)
            productField.IsAscendSort = false;

            // Specify which data field to sort by (0 = first data field, i.e., "Sales")
            productField.AutoSortField = 0;

            // Refresh the pivot table data and calculate the results
            pivotTable.RefreshDataFlag = true;
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotFieldDescendingCustomSort.xlsx");
        }
    }
}