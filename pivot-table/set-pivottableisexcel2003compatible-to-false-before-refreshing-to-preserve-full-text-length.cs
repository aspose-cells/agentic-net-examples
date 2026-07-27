using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and add sample data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            dataSheet.Cells["A1"].Value = "Product";
            dataSheet.Cells["B1"].Value = "Description";
            dataSheet.Cells["C1"].Value = "Quantity";

            dataSheet.Cells["A2"].Value = "Item1";
            dataSheet.Cells["B2"].Value = "Short description";
            dataSheet.Cells["C2"].Value = 10;

            dataSheet.Cells["A3"].Value = "Item2";
            // Very long text that exceeds 255 characters
            dataSheet.Cells["B3"].Value = new string('X', 300);
            dataSheet.Cells["C3"].Value = 20;

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table (source range A1:C3, destination at A4)
            int pivotIndex = pivotSheet.PivotTables.Add("A1:C3", "A4", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure fields: Product as row, Description as data (count)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);     // Description

            // Disable Excel 2003 compatibility to keep full text length
            pivotTable.IsExcel2003Compatible = false;

            // Refresh data and calculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_Excel2003Compatibility.xlsx");
        }
    }
}