// Title: C# – Scan Excel workbook for INDIRECT formulas and evaluate performance with Aspose.Cells
// Description: Loads an Excel file, enables the calculation chain, runs an initial calculation, then walks every used cell in each worksheet. Formulas containing the volatile INDIRECT function are collected, displayed with sheet name and address, and a note explains how INDIRECT forces full recalculations that can degrade speed in large workbooks.
// Keywords: Aspose.Cells INDIRECT scan C# | detect volatile Excel functions .NET | list INDIRECT formulas programmatically | Excel calculation performance INDIRECT | C# workbook formula analysis
// Common Searches: How to find INDIRECT formulas using Aspose.Cells | Identify volatile functions in a .NET Excel workbook | Performance impact of INDIRECT in large spreadsheets | C# code to list cells with INDIRECT
// Developer Intent: Locate every formula that uses INDIRECT and understand its effect on workbook recalculation speed.
// Use Cases: Create an audit report of all INDIRECT formulas before optimizing a workbook. | Prioritize refactoring of volatile functions to improve calculation time in enterprise spreadsheets. | Export cells with INDIRECT to a CSV for review by business analysts.
// AI Prompts: Generate C# Aspose.Cells code that extracts all INDIRECT formulas and writes them to a CSV file. | Explain how disabling the calculation chain influences scanning speed for volatile functions in Aspose.Cells. | Suggest alternatives to INDIRECT for faster recalculation in large Excel workbooks.

using System;
using Aspose.Cells;
using System.Collections.Generic;

// Loads an Excel file, enables the calculation chain, runs an initial calculation, then walks every used cell in each worksheet. Formulas containing the volatile INDIRECT function are collected, displayed with sheet name and address, and a note explains how INDIRECT forces full recalculations that can degrade speed in large workbooks.
class IndirectFormulaScanner
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable calculation chain to improve subsequent dependent queries (optional)
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Perform an initial calculation so that all formulas are parsed
        workbook.CalculateFormula();

        // Collect formulas that contain the INDIRECT function
        List<string> indirectFormulas = new List<string>();

        // Iterate through every worksheet and its used cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
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
                        if (!string.IsNullOrEmpty(formula) &&
                            formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            indirectFormulas.Add($"{sheet.Name}!{cell.Name}: {formula}");
                        }
                    }
                }
            }
        }

        // Output the list of INDIRECT formulas
        Console.WriteLine("Formulas using INDIRECT:");
        foreach (string entry in indirectFormulas)
        {
            Console.WriteLine(entry);
        }

        // Explain the performance impact of INDIRECT
        Console.WriteLine("\nPerformance note: INDIRECT is a volatile function. Any change in the workbook forces a full recalculation, which can significantly slow down calculation on large workbooks.");

        // Save the workbook (no changes made, but required by lifecycle rules)
        workbook.Save("output.xlsx");
    }
}
