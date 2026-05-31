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

            // ----- Set up header row -----
            sheet.Cells["A1"].PutValue("Quantity");
            sheet.Cells["B1"].PutValue("UnitPrice");
            sheet.Cells["C1"].PutValue("Total");

            // ----- Add sample data rows -----
            // Row 2
            sheet.Cells["A2"].PutValue(5);
            sheet.Cells["B2"].PutValue(12);
            // Row 3
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["B3"].PutValue(7);
            // Row 4
            sheet.Cells["A4"].PutValue(8);
            sheet.Cells["B4"].PutValue(15);

            // ----- Create a table that spans the data (including header) -----
            // The range A1:C4 covers the header and three data rows
            int tableIndex = sheet.ListObjects.Add("A1", "C4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // ----- Insert formula for each data row -----
            // The Total column is the third column in the table (zero‑based index 2)
            // Data rows start at offset 1 (offset 0 is the header)
            int dataRowCount = 3; // we added three data rows
            for (int i = 0; i < dataRowCount; i++)
            {
                // Structured reference "@[Quantity]" and "@[UnitPrice]" refer to the current row's cells
                table.PutCellFormula(rowOffset: i + 1, columnOffset: 2,
                    formula: "=[@Quantity]*[@UnitPrice]");
            }

            // Save the workbook
            workbook.Save("QuantityUnitPriceTotal.xlsx");
        }
    }
}