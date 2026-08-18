// Title: Extract formula dependency graph after Workbook.CalculateFormula in C# with Aspose.Cells
// Description: Demonstrates how to enable the calculation chain, run Workbook.CalculateFormula, and then enumerate each formula cell to list its precedents (using GetPrecedentsInCalculation) and its recursive dependents (using GetDependentsInCalculation). The example logs the relationships to the console and saves the workbook, providing a practical way to debug and analyze formula dependencies in .NET.
// Keywords: Aspose.Cells | C# | Workbook.CalculateFormula | formula precedents | formula dependents | calculation chain | dependency graph | GetPrecedentsInCalculation | GetDependentsInCalculation | Excel formula debugging | cell dependency extraction
// Common Searches: Aspose.Cells get formula precedents after calculation | How to list dependent cells with Aspose.Cells .NET | Enable calculation chain for dependency tracking Aspose | Debug Excel formula graph using Aspose.Cells | Retrieve recursive dependents in Aspose.Cells
// Developer Intent: Retrieve and log the full formula dependency graph (precedents and dependents) after workbook calculation.
// Use Cases: Perform impact analysis to see which cells affect a specific formula. | Detect circular references by examining recursive dependent chains. | Create a textual or visual map of calculation order for troubleshooting complex spreadsheets. | Export dependency information for audit or documentation purposes.
// AI Prompts: Generate C# code that returns a Dictionary<string, List<string>> where each key is a formula cell and the value is its list of precedent cell names using Aspose.Cells. | Write a routine that writes the complete dependency graph (both precedents and dependents) to a JSON file after calling Workbook.CalculateFormula. | Explain the role of EnableCalculationChain in influencing GetPrecedentsInCalculation and GetDependentsInCalculation results.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsDependencyGraphDemo
{
    // Demonstrates how to enable the calculation chain, run Workbook.CalculateFormula, and then enumerate each formula cell to list its precedents (using GetPrecedentsInCalculation) and its recursive dependents (using GetDependentsInCalculation). The example logs the relationships to the console and saves the workbook, providing a practical way to debug and analyze formula dependencies in .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data and formulas to build a dependency graph
            cells["C1"].PutValue(10);                 // Source value
            cells["B1"].Formula = "C1*2";             // B1 depends on C1
            cells["A1"].Formula = "B1+5";             // A1 depends on B1 (and indirectly on C1)
            cells["D1"].Formula = "A1+B1";            // D1 depends on A1 and B1

            // Enable calculation chain to allow dependency tracking
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Iterate through all used cells to log their dependencies
            foreach (Cell cell in cells)
            {
                // Process only formula cells
                if (!cell.IsFormula) continue;

                Console.WriteLine($"Cell {cell.Name} (Formula: {cell.Formula})");

                // Get precedents (cells this cell depends on) during calculation
                IEnumerator precedentsEnum = cell.GetPrecedentsInCalculation();
                if (precedentsEnum != null)
                {
                    Console.WriteLine("  Precedents:");
                    while (precedentsEnum.MoveNext())
                    {
                        // Each item is a ReferredArea representing a referenced range
                        ReferredArea area = (ReferredArea)precedentsEnum.Current;
                        // For single-cell references, display the cell name
                        string refName = area.IsArea
                            ? $"{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}:{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}"
                            : CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                        Console.WriteLine($"    - {refName}");
                    }
                }

                // Get dependents (cells whose calculation result depends on this cell)
                IEnumerator dependentsEnum = cell.GetDependentsInCalculation(true);
                if (dependentsEnum != null)
                {
                    Console.WriteLine("  Dependents (recursive):");
                    while (dependentsEnum.MoveNext())
                    {
                        if (dependentsEnum.Current is Cell dependentCell)
                        {
                            Console.WriteLine($"    - {dependentCell.Name}");
                        }
                    }
                }

                Console.WriteLine(); // Blank line for readability
            }

            // Save the workbook (using the standard save rule)
            workbook.Save("DependencyGraphDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
