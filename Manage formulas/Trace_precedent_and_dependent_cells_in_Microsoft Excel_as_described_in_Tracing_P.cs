using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsPrecedentDependentDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Enable calculation chain to allow tracing of precedents/dependents in calculation
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Calculate all formulas so that the calculation chain is built
            workbook.CalculateFormula();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Trace calculation precedents of cell A2
            // -------------------------------------------------
            Cell targetCell = cells["A2"];
            IEnumerator precedentsEnum = targetCell.GetPrecedentsInCalculation();

            Console.WriteLine("Precedents (in calculation) of cell A2:");
            if (precedentsEnum != null)
            {
                while (precedentsEnum.MoveNext())
                {
                    // Each item is a ReferredArea describing a referenced range
                    ReferredArea area = (ReferredArea)precedentsEnum.Current;
                    string sheetName = area.SheetName ?? sheet.Name;
                    string start = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    string end = area.IsArea ? ":" + CellsHelper.CellIndexToName(area.EndRow, area.EndColumn) : string.Empty;
                    Console.WriteLine($"{sheetName}!{start}{end}");
                }
            }
            else
            {
                Console.WriteLine("No calculation precedents found (cell may not contain a formula).");
            }

            // -------------------------------------------------
            // 2. Trace calculation dependents of cell B1 (recursive)
            // -------------------------------------------------
            IEnumerator dependentsEnum = cells.GetDependentsInCalculation(0, 1, true); // row 0, column 1 => B1

            Console.WriteLine("\nDependents (in calculation, recursive) of cell B1:");
            if (dependentsEnum != null)
            {
                while (dependentsEnum.MoveNext())
                {
                    Cell depCell = (Cell)dependentsEnum.Current;
                    Console.WriteLine(depCell.Name);
                }
            }
            else
            {
                Console.WriteLine("No calculation dependents found.");
            }

            // -------------------------------------------------
            // 3. Non‑calculation precedents of cell C3 (all references in formula)
            // -------------------------------------------------
            ReferredAreaCollection allPrecedents = cells["C3"].GetPrecedents();

            Console.WriteLine("\nAll precedents (including those not used in calculation) of cell C3:");
            if (allPrecedents != null)
            {
                foreach (ReferredArea area in allPrecedents)
                {
                    string sheetName = area.SheetName ?? sheet.Name;
                    string start = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    string end = area.IsArea ? ":" + CellsHelper.CellIndexToName(area.EndRow, area.EndColumn) : string.Empty;
                    Console.WriteLine($"{sheetName}!{start}{end}");
                }
            }
            else
            {
                Console.WriteLine("Cell C3 does not contain a formula.");
            }

            // -------------------------------------------------
            // 4. Non‑calculation dependents of cell D4 (direct references only)
            // -------------------------------------------------
            Cell[] directDependents = cells.GetDependents(false, 3, 3); // row 3, column 3 => D4

            Console.WriteLine("\nDirect dependents of cell D4:");
            if (directDependents != null && directDependents.Length > 0)
            {
                foreach (Cell dep in directDependents)
                {
                    Console.WriteLine(dep.Name);
                }
            }
            else
            {
                Console.WriteLine("No direct dependents found.");
            }

            // Save the workbook (optional – can be used to verify that no changes broke the file)
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}