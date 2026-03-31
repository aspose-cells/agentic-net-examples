using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsPrecedentsDependentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Enable calculation chain to allow precedent/dependent queries
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Ensure all formulas are calculated so the dependency graph is built
            workbook.CalculateFormula();

            // Choose a cell to analyze (e.g., B1)
            Cell targetCell = workbook.Worksheets[0].Cells["B1"];

            // -------------------------
            // 1. Get precedents used by the cell during calculation
            // -------------------------
            Console.WriteLine($"Precedents of {targetCell.Name}:");
            IEnumerator precedentsEnum = targetCell.GetPrecedentsInCalculation();
            if (precedentsEnum != null)
            {
                while (precedentsEnum.MoveNext())
                {
                    // Each item is a ReferredArea describing a precedent range
                    ReferredArea area = (ReferredArea)precedentsEnum.Current;
                    string start = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    string end = area.IsArea ? CellsHelper.CellIndexToName(area.EndRow, area.EndColumn) : start;
                    string sheet = string.IsNullOrEmpty(area.SheetName) ? targetCell.Worksheet.Name : area.SheetName;
                    Console.WriteLine($"- {sheet}!{start}{(area.IsArea ? $":{end}" : "")}");
                }
            }
            else
            {
                Console.WriteLine("No precedents found (cell may not contain a formula).");
            }

            // -------------------------
            // 2. Get dependents whose calculated result depends on the cell (recursive)
            // -------------------------
            Console.WriteLine($"\nDependents of {targetCell.Name} (recursive):");
            IEnumerator dependentsEnum = targetCell.GetDependentsInCalculation(true);
            if (dependentsEnum != null)
            {
                while (dependentsEnum.MoveNext())
                {
                    Cell depCell = (Cell)dependentsEnum.Current;
                    Console.WriteLine($"- {depCell.Name} (Value: {depCell.Value})");
                }
            }
            else
            {
                Console.WriteLine("No dependents found.");
            }

            // -------------------------
            // 3. Alternative way: use Cells.GetDependentsInCalculation(row, column, recursive)
            // -------------------------
            Console.WriteLine($"\nDependents of {targetCell.Name} via Cells method:");
            IEnumerator dependentsEnum2 = workbook.Worksheets[0].Cells.GetDependentsInCalculation(
                targetCell.Row, targetCell.Column, true);
            if (dependentsEnum2 != null)
            {
                while (dependentsEnum2.MoveNext())
                {
                    Cell depCell = (Cell)dependentsEnum2.Current;
                    Console.WriteLine($"- {depCell.Name}");
                }
            }

            // Save the workbook (optional, e.g., after modifications)
            workbook.Save("Output.xlsx");
        }
    }
}