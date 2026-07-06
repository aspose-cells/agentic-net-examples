using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowAndRecalculate
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some initial data and a formula that depends on the data
            cells["A1"].PutValue(10);               // Row 0
            cells["A2"].PutValue(20);               // Row 1 (will be shifted after insertion)
            cells["A3"].Formula = "=A1+A2";         // Depends on A1 and A2

            // Insert a new row at index 1 (between the original rows 0 and 1)
            // This uses the Cells.InsertRow method rule
            cells.InsertRow(1);

            // After inserting the row, recalculate all formulas so that dependent
            // cells (e.g., A3) reflect the new layout.
            // This uses the Workbook.CalculateFormula method rule.
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("InsertedRowWithRecalculatedFormulas.xlsx");
        }
    }
}