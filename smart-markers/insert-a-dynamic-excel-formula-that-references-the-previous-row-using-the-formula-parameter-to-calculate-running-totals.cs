using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsRunningTotalDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Set up table headers
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("RunningTotal");

            // -------------------------------------------------
            // Add sample data rows (ID and Value)
            // -------------------------------------------------
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);   // RunningTotal for first row = Value

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(30);

            // -------------------------------------------------
            // Create a ListObject (table) that covers the data range
            // -------------------------------------------------
            // Table range: from A1 to C4 (including header row)
            int tableIndex = sheet.ListObjects.Add("A1", "C4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // -------------------------------------------------
            // Set RunningTotal for the first data row (row offset 1, column offset 2)
            // -------------------------------------------------
            // Row offset is zero‑based within the table (0 = header row)
            table.PutCellValue(1, 2, 10); // C2 = 10 (same as first Value)

            // -------------------------------------------------
            // Insert running‑total formulas for subsequent rows
            // -------------------------------------------------
            // Row offset 2 corresponds to worksheet row 3 (second data row)
            // Formula: =C2+B3   (previous RunningTotal + current Value)
            table.PutCellFormula(2, 2, "=C2+B3");

            // Row offset 3 corresponds to worksheet row 4 (third data row)
            // Formula: =C3+B4
            table.PutCellFormula(3, 2, "=C3+B4");

            // -------------------------------------------------
            // Calculate all formulas in the workbook
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("RunningTotalTable.xlsx");
        }
    }
}