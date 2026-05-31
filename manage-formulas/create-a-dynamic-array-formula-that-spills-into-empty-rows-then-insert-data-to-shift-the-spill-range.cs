using System;
using Aspose.Cells;

namespace DynamicArraySpillDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Prepare data that will drive the dynamic array formula
            //    B1 will hold the number of rows the SEQUENCE function should generate
            cells["B1"].PutValue(3); // initial spill size = 3 rows

            // 3. Set a dynamic array formula in A1 that spills vertically based on B1
            //    The formula will produce a sequence 1..B1 in column A
            Cell startCell = cells["A1"];
            startCell.SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

            // 4. Calculate formulas and refresh dynamic array spill range
            workbook.CalculateFormula();
            workbook.RefreshDynamicArrayFormulas(true);

            // 5. Insert two rows below the header (row index 1) to shift the spill range down
            //    This will move the existing spilled values down, creating empty rows above them
            cells.InsertRows(1, 2, true); // insert 2 rows at row 2 (zero‑based index 1)

            // 6. Update the driver cell to expand the spill range
            cells["B1"].PutValue(5); // now we expect 5 rows in the spill

            // 7. Refresh dynamic array formulas again so the spill expands into the newly inserted rows
            workbook.CalculateFormula();
            workbook.RefreshDynamicArrayFormulas(true);

            // 8. Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("DynamicArraySpillShifted.xlsx");
        }
    }
}