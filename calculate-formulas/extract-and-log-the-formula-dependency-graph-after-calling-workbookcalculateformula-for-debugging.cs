using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsDependencyGraph
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data and formulas to create a dependency chain
            cells["C1"].PutValue(10);               // Source value
            cells["B1"].Formula = "C1*2";           // Depends on C1
            cells["A1"].Formula = "B1+5";           // Depends on B1 (and indirectly on C1)

            // Enable calculation chain and calculate formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Iterate through all used cells to build and log the dependency graph
            Console.WriteLine("Formula Dependency Graph:");
            foreach (Cell cell in cells)
            {
                // Process only formula cells
                if (!cell.IsFormula) continue;

                string cellName = cell.Name;
                int row = cell.Row;
                int column = cell.Column;

                // Log precedents (cells this cell depends on)
                IEnumerator precedentsEnum = cell.GetPrecedentsInCalculation();
                if (precedentsEnum != null)
                {
                    Console.Write($"{cellName} depends on: ");
                    bool first = true;
                    while (precedentsEnum.MoveNext())
                    {
                        if (precedentsEnum.Current is ReferredArea area)
                        {
                            // For simplicity, list the start cell of each referred area
                            string refName = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                            if (!first) Console.Write(", ");
                            Console.Write(refName);
                            first = false;
                        }
                    }
                    Console.WriteLine();
                }

                // Log dependents (cells that depend on this cell)
                IEnumerator dependentsEnum = cells.GetDependentsInCalculation(row, column, true);
                if (dependentsEnum != null)
                {
                    Console.Write($"{cellName} is depended on by: ");
                    bool first = true;
                    while (dependentsEnum.MoveNext())
                    {
                        if (dependentsEnum.Current is Cell depCell)
                        {
                            if (!first) Console.Write(", ");
                            Console.Write(depCell.Name);
                            first = false;
                        }
                    }
                    Console.WriteLine();
                }

                Console.WriteLine(); // Blank line for readability
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("DependencyGraphDemo.xlsx");
        }
    }
}