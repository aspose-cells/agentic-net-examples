using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsTraceDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain to allow tracing of precedents/dependents
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Perform full calculation so the chain is built
            workbook.CalculateFormula();

            // Choose the cell to trace (example: cell B2)
            Cell targetCell = workbook.Worksheets[0].Cells["B2"];

            // -------------------- Trace Precedents (used in calculation) --------------------
            IEnumerator precedentsEnum = targetCell.GetPrecedentsInCalculation();

            Console.WriteLine($"Precedents of {targetCell.Name} (used in calculation):");
            if (precedentsEnum != null)
            {
                while (precedentsEnum.MoveNext())
                {
                    // Each item is a ReferredArea describing a referenced range or cell
                    ReferredArea area = (ReferredArea)precedentsEnum.Current;
                    string areaDesc = area.IsArea
                        ? $"{area.SheetName}!{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}:{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}"
                        : $"{area.SheetName}!{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}";

                    if (area.IsExternalLink)
                        areaDesc = $"[{area.ExternalFileName}]{areaDesc}";

                    Console.WriteLine($"  {areaDesc}");
                }
            }
            else
            {
                Console.WriteLine("  No precedents found (cell may not contain a formula).");
            }

            // -------------------- Trace Dependents (cells whose result depends on target) --------------------
            // Recursive = true to include indirect dependents
            IEnumerator dependentsEnum = workbook.Worksheets[0].Cells.GetDependentsInCalculation(targetCell.Row, targetCell.Column, true);

            Console.WriteLine($"\nDependents of {targetCell.Name} (calculation chain, recursive):");
            if (dependentsEnum != null)
            {
                while (dependentsEnum.MoveNext())
                {
                    // Each item is a Cell object
                    Cell dependentCell = (Cell)dependentsEnum.Current;
                    Console.WriteLine($"  {dependentCell.Name} (Formula: {dependentCell.Formula})");
                }
            }
            else
            {
                Console.WriteLine("  No dependents found.");
            }

            // Save the workbook after calculation (optional)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}