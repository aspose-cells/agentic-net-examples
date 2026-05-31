using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a formula in cell E3 that references empty cells, so its result will be zero
        worksheet.Cells["E3"].Formula = "=A1+A2";

        // Calculate formulas to obtain the result
        workbook.CalculateFormula();

        // Output the unchanged formula stored in the cell
        Console.WriteLine("Formula in E3: " + worksheet.Cells["E3"].Formula);

        // Output the calculated value (zero) of the cell
        Console.WriteLine("Value in E3: " + worksheet.Cells["E3"].Value);
    }
}