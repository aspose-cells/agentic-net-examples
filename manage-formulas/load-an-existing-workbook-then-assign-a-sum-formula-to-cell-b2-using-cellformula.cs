using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Assign a SUM formula to cell B2 (summing A1 through A5)
        worksheet.Cells["B2"].Formula = "=SUM(A1:A5)";

        // Recalculate formulas so the result is stored in the cell
        workbook.CalculateFormula();

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}