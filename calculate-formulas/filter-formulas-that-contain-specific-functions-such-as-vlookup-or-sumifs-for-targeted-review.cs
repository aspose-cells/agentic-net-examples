// Title: C# Example: Filter and Highlight Cells Containing VLOOKUP or SUMIFS with Aspose.Cells
// Description: Loads an Excel workbook, scans the used range of the first worksheet, and collects every cell whose formula includes VLOOKUP or SUMIFS (case‑insensitive). The script prints each address and formula, highlights the matching cells in yellow, and saves the workbook with the visual markers.
// Keywords: Aspose.Cells | C# | filter formulas | VLOOKUP detection | SUMIFS detection | find formulas by function | highlight cells | Excel automation .NET | used range scan | formula search example | GitHub code sample
// Common Searches: Aspose.Cells find VLOOKUP formulas C# | highlight SUMIFS cells using Aspose.Cells | search Excel formulas by function name .NET | filter cells containing specific functions Aspose.Cells | C# code to locate VLOOKUP or SUMIFS in workbook
// Developer Intent: Locate and visually flag every cell whose formula contains VLOOKUP or SUMIFS.
// Use Cases: Audit a workbook before migration by listing all VLOOKUP and SUMIFS formulas. | Create a visual report that highlights performance‑critical SUMIFS calculations. | Assist formula reviewers by marking cells that rely on VLOOKUP for easier validation.
// AI Prompts: Generate a C# Aspose.Cells snippet that extracts cell addresses with the INDEX function and writes them to a CSV file. | Adapt the provided script to export matched cell names and formulas to a JSON document instead of console output. | Explain how to load target function names from an external configuration file and apply them in the formula‑filtering loop.

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

// Loads an Excel workbook, scans the used range of the first worksheet, and collects every cell whose formula includes VLOOKUP or SUMIFS (case‑insensitive). The script prints each address and formula, highlights the matching cells in yellow, and saves the workbook with the visual markers.
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

        // Scan the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
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
        }

        // Report the findings
        Console.WriteLine($"Found {matchedCells.Count} cells containing VLOOKUP or SUMIFS:");
        foreach (Cell cell in matchedCells)
        {
            Console.WriteLine($"{cell.Name}: {cell.Formula}");
        }

        // Optional: highlight the matched cells for visual review
        foreach (Cell cell in matchedCells)
        {
            Style style = cell.GetStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            cell.SetStyle(style);
        }

        // Save the workbook with highlights (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
