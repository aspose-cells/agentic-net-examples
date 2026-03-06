using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsPrecedentDependentDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Enable calculation chain so that precedent/dependent tracing works
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Calculate all formulas to build the calculation chain
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Example 1: Get precedents (cells that influence a formula) for cell A2
            // -------------------------------------------------
            Cell targetCell = cells["A2"]; // Change as needed
            IEnumerator precedentsEnum = targetCell.GetPrecedentsInCalculation();

            Console.WriteLine($"Precedents used by {targetCell.Name}:");
            if (precedentsEnum != null)
            {
                while (precedentsEnum.MoveNext())
                {
                    // Each item is a ReferredArea describing a referenced range
                    ReferredArea area = (ReferredArea)precedentsEnum.Current;
                    // Build a readable address string
                    string address = $"{area.SheetName}!{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}";
                    if (area.IsArea)
                    {
                        address += $":{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}";
                    }
                    Console.WriteLine($"  {address}");
                }
            }
            else
            {
                Console.WriteLine("  No precedents (cell may not contain a formula).");
            }

            // -------------------------------------------------
            // Example 2: Get dependents (cells whose results depend on a specific cell) for cell B1
            // -------------------------------------------------
            int row = cells["B1"].Row;
            int column = cells["B1"].Column;
            bool recursive = true; // include indirect dependents

            IEnumerator dependentsEnum = cells.GetDependentsInCalculation(row, column, recursive);

            Console.WriteLine($"\nDependents of {cells["B1"].Name}:");
            if (dependentsEnum != null)
            {
                while (dependentsEnum.MoveNext())
                {
                    // Each item is a Cell object
                    Cell dependentCell = (Cell)dependentsEnum.Current;
                    Console.WriteLine($"  {dependentCell.Name}");
                }
            }
            else
            {
                Console.WriteLine("  No dependents found.");
            }

            // -------------------------------------------------
            // Example 3: Using Cell.GetDependentsInCalculation(bool) directly
            // -------------------------------------------------
            IEnumerator dependentsEnum2 = cells["C1"].GetDependentsInCalculation(recursive);
            Console.WriteLine($"\nDependents of {cells["C1"].Name} (using Cell method):");
            if (dependentsEnum2 != null)
            {
                while (dependentsEnum2.MoveNext())
                {
                    Cell dep = (Cell)dependentsEnum2.Current;
                    Console.WriteLine($"  {dep.Name}");
                }
            }
            else
            {
                Console.WriteLine("  No dependents found.");
            }

            // Save the workbook (optional, demonstrates lifecycle rule)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}