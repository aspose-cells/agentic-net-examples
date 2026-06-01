using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Calculate all formulas in the loaded workbook
        workbook.CalculateFormula();

        // Save the workbook after calculation
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}