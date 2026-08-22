// Title: Extract and log the Excel formula dependency graph after calling Workbook.CalculateFormula with Aspose.Cells for .NET
// AI Prompts: Write C# code that enables the calculation chain, runs Workbook.CalculateFormula, and prints each formula cell’s precedents and dependents using Aspose.Cells. | Provide a step‑by‑step example that creates inter‑dependent formulas, calculates them, and outputs the complete formula dependency graph for debugging.
// Common Searches: Aspose.Cells get precedent cells for a formula after CalculateFormula | How to retrieve dependent cells of a formula in Aspose.Cells .NET | Enable calculation chain to access formula dependency graph with Aspose.Cells | Debug formula dependencies in an Excel workbook using Aspose.Cells C# | Aspose.Cells log full formula dependency chain after workbook calculation
// Tags: Aspose.Cells enable calculation chain | Aspose.Cells get formula precedents C# | Aspose.Cells retrieve formula dependents .NET | Aspose.Cells log formula dependency graph | Aspose.Cells Workbook.CalculateFormula debugging

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, turns on the calculation chain, defines several inter‑related formulas, runs Workbook.CalculateFormula, and then iterates over each formula cell to display its name, formula, referenced precedents, and dependent cells before saving the file.
    public class FormulaDependencyGraphDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Enable calculation chain so dependency methods work
                workbook.Settings.FormulaSettings.EnableCalculationChain = true;

                // Set up sample data and formulas to create a dependency graph
                cells["C1"].PutValue(10);               // source value
                cells["B1"].Formula = "C1*2";           // depends on C1
                cells["A1"].Formula = "B1+5";           // depends on B1
                cells["D1"].Formula = "A1+B1";          // depends on A1 and B1
                cells["E1"].Formula = "SUM(A1:D1)";     // depends on A1,B1,C1,D1

                // Calculate formulas to build the calculation chain
                workbook.CalculateFormula();

                // Iterate over all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula)
                    {
                        Console.WriteLine($"Cell {cell.Name} (Formula: {cell.Formula})");

                        // Get precedents (cells referenced by this formula)
                        IEnumerator precedents = cell.GetPrecedentsInCalculation();
                        if (precedents != null)
                        {
                            Console.Write("  Precedents: ");
                            bool first = true;
                            while (precedents.MoveNext())
                            {
                                ReferredArea area = (ReferredArea)precedents.Current;
                                string refName = area.IsArea
                                    ? $"{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}:{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}"
                                    : CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                                if (!first) Console.Write(", ");
                                Console.Write(refName);
                                first = false;
                            }
                            Console.WriteLine();
                        }

                        // Get dependents (cells whose calculation result depends on this cell)
                        IEnumerator dependents = cell.GetDependentsInCalculation(true);
                        if (dependents != null)
                        {
                            Console.Write("  Dependents: ");
                            bool firstDep = true;
                            while (dependents.MoveNext())
                            {
                                if (dependents.Current is Cell depCell)
                                {
                                    if (!firstDep) Console.Write(", ");
                                    Console.Write(depCell.Name);
                                    firstDep = false;
                                }
                            }
                            Console.WriteLine();
                        }

                        Console.WriteLine();
                    }
                }

                // Save the workbook (lifecycle rule)
                workbook.Save("FormulaDependencyGraphDemo.xlsx");
                Console.WriteLine("Workbook saved as FormulaDependencyGraphDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
