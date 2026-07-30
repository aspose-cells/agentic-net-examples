// Title: Extract and Log Formula Dependency Graph (Precedents & Dependents) after Workbook.CalculateFormula – Aspose.Cells .NET Example
// Description: Demonstrates how to enable the calculation chain, run Workbook.CalculateFormula, and then iterate all formula cells to retrieve their precedents with GetPrecedentsInCalculation and their dependents with GetDependentsInCalculation(true). The code prints a clear dependency report to the console and optionally saves the workbook.
// Keywords: Aspose.Cells formula precedents | Aspose.Cells dependency graph | GetPrecedentsInCalculation C# | GetDependentsInCalculation example | EnableCalculationChain Aspose.Cells | debug spreadsheet formulas .NET | Aspose.Cells CalculateFormula logging | C# Excel formula dependency
// Common Searches: how to get precedent cells after calculating formulas Aspose.Cells | list dependent cells for a formula using Aspose.Cells .NET | Aspose.Cells extract formula dependency chain | debug formula order with Aspose.Cells GetPrecedentsInCalculation | Aspose.Cells GetDependentsInCalculation multi‑sheet example
// Developer Intent: The developer needs to extract and log the full formula dependency graph—both direct precedents and dependents—after invoking Workbook.CalculateFormula in an Aspose.Cells .NET workbook.
// Use Cases: Display direct precedents for each formula cell to troubleshoot complex calculations. | Identify all cells that rely on a specific source cell, helping detect circular references or impact analysis. | Generate a textual report of cell relationships for auditing or documentation across one or multiple worksheets.
// AI Prompts: Show how to recursively collect indirect precedents for each formula cell in Aspose.Cells. | Provide a method that returns a dictionary mapping each formula cell to its list of precedents and dependents. | Explain the behavior of GetPrecedentsInCalculation and GetDependentsInCalculation when formulas span multiple worksheets.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsDependencyGraphDemo
{
    // Demonstrates how to enable the calculation chain, run Workbook.CalculateFormula, and then iterate all formula cells to retrieve their precedents with GetPrecedentsInCalculation and their dependents with GetDependentsInCalculation(true). The code prints a clear dependency report to the console and optionally saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data and formulas to build a dependency chain
            cells["C1"].PutValue(10);                 // Source value
            cells["B1"].Formula = "C1*2";             // B1 depends on C1
            cells["A1"].Formula = "B1+5";             // A1 depends on B1 (and indirectly on C1)
            cells["D1"].Formula = "A1+B1";            // D1 depends on A1 and B1

            // Enable calculation chain and calculate all formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Iterate through all used cells to build the dependency graph
            foreach (Cell cell in cells)
            {
                // Process only formula cells
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    Console.WriteLine($"Cell {cell.Name} (Formula: {cell.Formula})");

                    // Get precedents (cells referenced by this formula during calculation)
                    IEnumerator precedentsEnum = cell.GetPrecedentsInCalculation();
                    if (precedentsEnum != null)
                    {
                        Console.WriteLine("  Precedents:");
                        while (precedentsEnum.MoveNext())
                        {
                            // Each item is a ReferredArea describing a referenced range
                            ReferredArea area = (ReferredArea)precedentsEnum.Current;
                            // For simplicity, log the start cell of the area
                            string startCell = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                            if (area.IsArea)
                            {
                                string endCell = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                                Console.WriteLine($"    {area.SheetName}!{startCell}:{endCell}");
                            }
                            else
                            {
                                Console.WriteLine($"    {area.SheetName}!{startCell}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("  No precedents.");
                    }

                    // Get dependents (cells whose calculated result depends on this cell)
                    IEnumerator dependentsEnum = cell.GetDependentsInCalculation(true);
                    if (dependentsEnum != null)
                    {
                        Console.WriteLine("  Dependents:");
                        while (dependentsEnum.MoveNext())
                        {
                            if (dependentsEnum.Current is Cell dependentCell)
                            {
                                Console.WriteLine($"    {dependentCell.Name}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("  No dependents.");
                    }

                    Console.WriteLine(); // Blank line for readability
                }
            }

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("DependencyGraphDemo.xlsx");
        }
    }
}
