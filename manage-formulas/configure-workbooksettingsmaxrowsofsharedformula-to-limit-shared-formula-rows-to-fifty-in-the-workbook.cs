using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Limit the maximum number of rows that a shared formula can span to 50
        workbook.Settings.MaxRowsOfSharedFormula = 50;

        // (Optional) Demonstrate setting a shared formula that respects the limit
        // This will populate the formula in rows B1 to B50
        workbook.Worksheets[0].Cells["B1"].SetSharedFormula("=A1", 50, 1);

        // Save the workbook to a file
        workbook.Save("LimitedSharedFormula.xlsx");
    }
}