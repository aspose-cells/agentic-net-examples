using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – calculate all formulas after loading a workbook
class Program
{
    static void Main()
    {
        // Load the workbook from a file.
        Workbook workbook = new Workbook("input.xlsx");

        // Calculate all formulas in the workbook.
        workbook.CalculateFormula();

        // Save the workbook after calculation.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}