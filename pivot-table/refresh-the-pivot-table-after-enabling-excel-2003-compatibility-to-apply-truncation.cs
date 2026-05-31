using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (will hold source data)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate source data with a long text (>255 characters) to demonstrate truncation
            string longText = new string('X', 300); // 300 characters
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Description");
            dataSheet.Cells["A2"].PutValue("Item1");
            dataSheet.Cells["B2"].PutValue(longText); // This will be truncated when Excel2003 compatibility is on

            // Add a second row with normal length text
            dataSheet.Cells["A3"].PutValue("Item2");
            dataSheet.Cells["B3"].PutValue("Short description");

            // Add a worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table based on the source data range
            int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B3", "A5", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields (Category as row, Description as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Description column

            // Ensure Excel 2003 compatibility is enabled (default is true, set explicitly)
            pivotTable.IsExcel2003Compatible = true;

            // Refresh the pivot table so that truncation is applied to long strings
            // Refresh all pivot tables in the workbook (could also use pivotSheet.RefreshPivotTables())
            workbook.Worksheets.RefreshPivotTables();

            // Save the workbook
            workbook.Save("PivotTable_Excel2003Compatibility.xlsx");
        }
    }
}