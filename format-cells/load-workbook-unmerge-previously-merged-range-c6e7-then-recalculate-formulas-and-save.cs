using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unmerge the previously merged range C6:E7
        // C6 -> row index 5, column index 2 (zero‑based)
        // Total rows = 2 (rows 6 and 7), total columns = 3 (C, D, E)
        worksheet.Cells.UnMerge(5, 2, 2, 3);

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}