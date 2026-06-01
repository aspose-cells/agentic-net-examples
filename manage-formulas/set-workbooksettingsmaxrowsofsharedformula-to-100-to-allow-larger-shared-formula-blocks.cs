using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Set the maximum number of rows that a shared formula can span
        // This allows shared formula blocks up to 100 rows
        workbook.Settings.MaxRowsOfSharedFormula = 100;

        // Example: set a shared formula that tries to cover 101 rows.
        // Because MaxRowsOfSharedFormula is 100, only the first 100 rows will receive the formula.
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        cells["B1"].SetSharedFormula("=A1", 101, 1);

        // Save the workbook (lifecycle: save)
        workbook.Save("SharedFormulaDemo.xlsx");
    }
}