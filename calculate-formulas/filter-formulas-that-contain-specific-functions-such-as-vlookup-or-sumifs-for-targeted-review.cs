using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

class FilterFormulas
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Functions we want to locate in formulas
        string[] targetFunctions = { "VLOOKUP", "SUMIFS" };

        // Store cells that contain any of the target functions
        List<Cell> matchedCells = new List<Cell>();

        // Scan all used cells in the worksheet
        foreach (Cell cell in cells)
        {
            if (cell.IsFormula)
            {
                string formula = cell.Formula;
                foreach (string func in targetFunctions)
                {
                    if (formula.IndexOf(func, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        matchedCells.Add(cell);
                        break; // No need to check other functions for this cell
                    }
                }
            }
        }

        // Output the addresses and formulas of matched cells
        Console.WriteLine("Cells containing VLOOKUP or SUMIFS:");
        foreach (Cell c in matchedCells)
        {
            Console.WriteLine($"{c.Name}: {c.Formula}");
        }

        // Optional: highlight matched cells for visual review
        foreach (Cell c in matchedCells)
        {
            Style style = c.GetStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            c.SetStyle(style);
        }

        // Save the workbook with highlights (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}