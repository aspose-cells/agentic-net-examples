using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the merged workbook from file
        Workbook workbook = new Workbook("merged.xlsx");

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook after calculation
        workbook.Save("merged_calculated.xlsx");
    }
}