using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet for source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate the worksheet with sample data, including a long text (>255 chars)
        dataSheet.Cells["A1"].Value = "Product";
        dataSheet.Cells["B1"].Value = "Description";

        dataSheet.Cells["A2"].Value = "Item1";
        // Long description to test Excel 2003 compatibility
        dataSheet.Cells["B2"].Value = new string('X', 300);

        // Add a new worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create a pivot table based on the data range A1:B2, place it starting at cell A4
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B2", "A4", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure the pivot table fields: Product as row, Description as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Description

        // Disable Excel 2003 compatibility to preserve full text length
        pivotTable.IsExcel2003Compatible = false;

        // Refresh the pivot table data and calculate the results
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook to a file
        workbook.Save("PivotExcel2003Compatibility.xlsx");
    }
}