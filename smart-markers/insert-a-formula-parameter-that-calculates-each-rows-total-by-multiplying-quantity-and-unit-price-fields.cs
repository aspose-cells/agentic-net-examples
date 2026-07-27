using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add headers for Quantity, UnitPrice and Total
            sheet.Cells["A1"].PutValue("Quantity");
            sheet.Cells["B1"].PutValue("UnitPrice");
            sheet.Cells["C1"].PutValue("Total");

            // Add some sample data rows
            sheet.Cells["A2"].PutValue(5);
            sheet.Cells["B2"].PutValue(12.5);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["B3"].PutValue(7.8);
            sheet.Cells["A4"].PutValue(10);
            sheet.Cells["B4"].PutValue(4.2);

            // Create a ListObject (table) that includes the data range (including header)
            // Parameters: first row, first column, last row, last column, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set the formula for the "Total" column to multiply Quantity and UnitPrice for each row
            // The column index for "Total" is 2 (zero‑based within the table)
            ListColumn totalColumn = table.ListColumns[2];
            // Use structured reference syntax; column names are enclosed in brackets
            totalColumn.Formula = "=[Quantity]*[UnitPrice]";

            // Save the workbook
            workbook.Save("QuantityUnitPriceTotal.xlsx");
        }
    }
}