using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Populate the first worksheet with sample data
        Worksheet dataSheet = wb.Worksheets[0];
        dataSheet.Name = "Data";
        dataSheet.Cells["A1"].Value = "Product";
        dataSheet.Cells["B1"].Value = "Description";
        dataSheet.Cells["C1"].Value = "Quantity";

        dataSheet.Cells["A2"].Value = "Prod1";
        // Long description to demonstrate Excel2003 truncation behavior
        dataSheet.Cells["B2"].Value = new string('X', 300);
        dataSheet.Cells["C2"].Value = 10;

        dataSheet.Cells["A3"].Value = "Prod2";
        dataSheet.Cells["B3"].Value = "Short description";
        dataSheet.Cells["C3"].Value = 20;

        // Add a new worksheet that will contain the pivot table
        Worksheet pivotSheet = wb.Worksheets.Add("Pivot");

        // Create a pivot table using the data range from the "Data" sheet
        // Parameters: source range, destination cell, pivot table name
        int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:C3", "A5", "PivotTable1");
        PivotTable pivot = pivotSheet.PivotTables[pivotIndex];

        // Configure pivot fields: Product as row, Quantity as data
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
        pivot.AddFieldToArea(PivotFieldType.Data, 2);  // Column index 2 -> Quantity

        // Enable Excel 2003 compatibility before refreshing the pivot table
        pivot.IsExcel2003Compatible = true;

        // Refresh the pivot cache and calculate the results
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook
        wb.Save("PivotExcel2003Compatible.xlsx");
    }
}