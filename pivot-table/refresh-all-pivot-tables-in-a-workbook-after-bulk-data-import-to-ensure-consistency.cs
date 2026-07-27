using System;
using Aspose.Cells;

class RefreshPivotTablesAfterImport
{
    static void Main()
    {
        // Load an existing workbook that contains one or more PivotTables.
        // Replace "SourceData.xlsx" with the path to your workbook.
        Workbook workbook = new Workbook("SourceData.xlsx");

        // -----------------------------------------------------------------
        // Bulk data import – example of updating a data range.
        // Adjust the range and values according to your actual import logic.
        // -----------------------------------------------------------------
        Worksheet dataSheet = workbook.Worksheets[0]; // assuming data is on the first sheet

        // Example: write header
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["B1"].PutValue("Sales");

        // Example: import 5 rows of data
        string[] products = { "Apple", "Orange", "Banana", "Grape", "Mango" };
        int[] sales = { 1200, 1500, 800, 950, 1100 };

        for (int i = 0; i < products.Length; i++)
        {
            dataSheet.Cells[i + 2, 0].PutValue(products[i]);   // Column A
            dataSheet.Cells[i + 2, 1].PutValue(sales[i]);     // Column B
        }

        // ---------------------------------------------------------------
        // Refresh all PivotTables in the workbook to reflect the new data.
        // Uses WorksheetCollection.RefreshPivotTables() method.
        // ---------------------------------------------------------------
        workbook.Worksheets.RefreshPivotTables();

        // Save the updated workbook.
        // Replace "RefreshedData.xlsx" with your desired output path.
        workbook.Save("RefreshedData.xlsx");
    }
}