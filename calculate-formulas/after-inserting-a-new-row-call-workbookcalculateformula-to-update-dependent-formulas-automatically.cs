using System;
using Aspose.Cells;

namespace AsposeCellsRowInsertAndRecalc
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some initial data
            // Row 0
            cells["A1"].PutValue(10);   // A1 = 10
            cells["B1"].PutValue(20);   // B1 = 20
            // Row 1
            cells["A2"].PutValue(30);   // A2 = 30
            cells["B2"].PutValue(40);   // B2 = 40

            // Set a formula that depends on the values above
            // C1 = SUM(A1:A2)  -> should be 40
            cells["C1"].Formula = "=SUM(A1:A2)";

            // Insert a new row at index 1 (between the two existing rows)
            // This will push the original row 1 (A2,B2) down to row 2
            cells.InsertRow(1); // lifecycle rule: insert row

            // After insertion, the formula in C1 still refers to A1:A2,
            // but now A2 is the newly inserted empty row.
            // Call CalculateFormula to recalculate dependent formulas.
            workbook.CalculateFormula(); // lifecycle rule: calculate formulas

            // Output the recalculated value of C1 to verify the update
            Console.WriteLine("Recalculated C1 value: " + cells["C1"].IntValue);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RowInsertedAndRecalculated.xlsx");
        }
    }
}