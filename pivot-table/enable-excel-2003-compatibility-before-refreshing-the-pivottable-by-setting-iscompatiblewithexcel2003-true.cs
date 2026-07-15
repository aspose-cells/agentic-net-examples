using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExcel2003Compatibility
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].Value = "Product";
            dataSheet.Cells["B1"].Value = "Description";
            dataSheet.Cells["C1"].Value = "Quantity";
            dataSheet.Cells["A2"].Value = "Item1";
            dataSheet.Cells["B2"].Value = new string('X', 300); // Long text to test 2003 limit
            dataSheet.Cells["C2"].Value = 10;
            dataSheet.Cells["A3"].Value = "Item2";
            dataSheet.Cells["B3"].Value = "Short desc";
            dataSheet.Cells["C3"].Value = 20;

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Add a pivot table based on the data range
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:C3", "A5", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);    // Quantity as data field

            // Enable Excel 2003 compatibility before refreshing
            pivotTable.IsExcel2003Compatible = true;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook (save rule)
            workbook.Save("PivotExcel2003Compatibility.xlsx");
        }
    }
}