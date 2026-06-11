using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert the SUM formula into cell G10 (range B2:B9)
        worksheet.Cells["G10"].Formula = "=SUM(B2:B9)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file (lifecycle rule: save)
        workbook.Save("output.xlsx");
    }
}