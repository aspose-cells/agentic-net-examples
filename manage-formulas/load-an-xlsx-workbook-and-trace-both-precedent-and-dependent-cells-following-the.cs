using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCalculationChainDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Enable calculation chain and calculate all formulas
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula();

            // Choose a cell to trace (example: A2)
            Cell targetCell = workbook.Worksheets[0].Cells["A2"];

            // Trace precedents used in calculation of the target cell
            IEnumerator precedentsEnum = targetCell.GetPrecedentsInCalculation();
            Console.WriteLine($"Precedents of {targetCell.Name}:");
            if (precedentsEnum != null)
            {
                while (precedentsEnum.MoveNext())
                {
                    // Each item is a ReferredArea describing a precedent range
                    ReferredArea area = (ReferredArea)precedentsEnum.Current;
                    Console.WriteLine($"- Sheet: {area.SheetName}, Range: {area.StartRow},{area.StartColumn} to {area.EndRow},{area.EndColumn}");
                }
            }
            else
            {
                Console.WriteLine("No precedents (cell may not contain a formula).");
            }

            // Trace dependents that rely on the target cell (recursive)
            IEnumerator dependentsEnum = targetCell.GetDependentsInCalculation(true);
            Console.WriteLine($"\nDependents of {targetCell.Name} (recursive):");
            if (dependentsEnum != null)
            {
                while (dependentsEnum.MoveNext())
                {
                    Cell depCell = (Cell)dependentsEnum.Current;
                    Console.WriteLine($"- {depCell.Name} (Formula: {depCell.Formula})");
                }
            }
            else
            {
                Console.WriteLine("No dependents found.");
            }

            // Optionally, save the workbook after any modifications
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to {outputPath}");
        }
    }
}