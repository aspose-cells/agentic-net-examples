using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExcel2003Compatibility
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

            dataSheet.Cells["A2"].Value = "Prod1";
            dataSheet.Cells["B2"].Value = "Short description";
            dataSheet.Cells["C2"].Value = 10;

            dataSheet.Cells["A3"].Value = "Prod2";
            dataSheet.Cells["B3"].Value = "Very long description that exceeds the 255‑character limit when used in an Excel 2003 pivot table. " +
                                          "This text is intentionally long to demonstrate the effect of the IsExcel2003Compatible flag.";
            dataSheet.Cells["C3"].Value = 20;

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table (source range A1:C3, destination start cell A5)
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:C3", "A5", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Quantity as data field

            // Enable Excel 2003 compatibility before refreshing
            pivotTable.IsExcel2003Compatible = true;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_Excel2003Compatible.xlsx");
        }
    }
}