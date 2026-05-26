using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable calculation chain to improve dependent analysis performance (optional)
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Collect formulas that contain the INDIRECT function
        List<string> indirectFormulas = new List<string>();

        // Scan every worksheet and every cell that has a formula
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            foreach (Cell cell in cells)
            {
                if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                    cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Store worksheet name, cell address and the formula text
                    indirectFormulas.Add($"{sheet.Name}!{cell.Name}: {cell.Formula}");
                }
            }
        }

        // Output the list of INDIRECT formulas
        Console.WriteLine("Formulas that use INDIRECT:");
        foreach (string entry in indirectFormulas)
        {
            Console.WriteLine(entry);
        }

        // Explain the performance impact
        Console.WriteLine();
        Console.WriteLine("Note: INDIRECT is a volatile function. It forces recalculation of all dependent cells whenever any cell changes,");
        Console.WriteLine("which can significantly slow down calculation speed in large workbooks.");

        // Recalculate the workbook (if needed) and save it
        workbook.CalculateFormula();
        workbook.Save("output.xlsx");
    }
}