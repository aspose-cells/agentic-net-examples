using System;
using System.Collections;
using Aspose.Cells;

namespace PrecedentsDependentsDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook.
            string inputPath = "Input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Enable calculation chain to allow tracing of precedents/dependents.
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Calculate all formulas so that the calculation chain is populated.
            workbook.CalculateFormula();

            // Select the cell to analyze (e.g., B1 on the first worksheet).
            Cell targetCell = workbook.Worksheets[0].Cells["B1"];

            // -------------------------------------------------
            // 1. Get all precedents appearing in the formula (including those not used in calculation).
            // -------------------------------------------------
            ReferredAreaCollection allPrecedents = targetCell.GetPrecedents();
            Console.WriteLine($"All precedents of {targetCell.Name}:");
            if (allPrecedents != null)
            {
                foreach (ReferredArea area in allPrecedents)
                {
                    Console.WriteLine(FormatReferredArea(area));
                }
            }
            else
            {
                Console.WriteLine("None (cell may not contain a formula).");
            }

            // -------------------------------------------------
            // 2. Get calculation precedents (only those that actually participated in the calculation).
            // -------------------------------------------------
            IEnumerator calcPrecedents = targetCell.GetPrecedentsInCalculation();
            Console.WriteLine($"\nCalculation precedents of {targetCell.Name}:");
            if (calcPrecedents != null)
            {
                while (calcPrecedents.MoveNext())
                {
                    ReferredArea area = (ReferredArea)calcPrecedents.Current;
                    Console.WriteLine(FormatReferredArea(area));
                }
            }
            else
            {
                Console.WriteLine("None (cell may not contain a formula or chain disabled).");
            }

            // -------------------------------------------------
            // 3. Get dependents whose calculated result depends on this cell (recursive).
            // -------------------------------------------------
            IEnumerator dependents = targetCell.GetDependentsInCalculation(true);
            Console.WriteLine($"\nRecursive dependents of {targetCell.Name}:");
            if (dependents != null)
            {
                while (dependents.MoveNext())
                {
                    Cell dep = (Cell)dependents.Current;
                    Console.WriteLine(dep.Name);
                }
            }
            else
            {
                Console.WriteLine("No dependents found.");
            }

            // Save the workbook (demonstrates the required save lifecycle).
            workbook.Save("Output.xlsx");
        }

        // Helper method to convert a ReferredArea into a readable string.
        static string FormatReferredArea(ReferredArea area)
        {
            string result = "";
            if (area.IsExternalLink)
            {
                result += $"[{area.ExternalFileName}]";
            }
            result += $"{area.SheetName}!";
            result += CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
            if (area.IsArea)
            {
                result += $":{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}";
            }
            return result;
        }
    }
}