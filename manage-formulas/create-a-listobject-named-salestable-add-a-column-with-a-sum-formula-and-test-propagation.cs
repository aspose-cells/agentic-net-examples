using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the table
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Item 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Item 2");
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["A4"].PutValue("Item 3");
        worksheet.Cells["B4"].PutValue(200);

        // Add a ListObject (table) that includes the data range
        // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
        int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
        ListObject salesTable = worksheet.ListObjects[tableIndex];

        // Assign a name to the table
        salesTable.DisplayName = "SalesTable";

        // Enable the totals row for the table
        salesTable.ShowTotals = true;

        // Set the totals calculation for the "Sales" column (index 1) to SUM
        salesTable.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

        // Add a new row to the table to test that the SUM total updates automatically
        // Row offset is zero‑based relative to the table (including header row)
        salesTable.PutCellValue(4, 0, "Item 4");   // Product column
        salesTable.PutCellValue(4, 1, 250);       // Sales column

        // Save the workbook to a file
        workbook.Save("SalesTableDemo.xlsx");
    }
}