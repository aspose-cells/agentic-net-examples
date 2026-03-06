using System;
using System.Collections;
using Aspose.Cells;

class TracePrecedentsDependents
{
    static void Main()
    {
        // Load the workbook from an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");

        // Enable the calculation chain so that precedent/dependent tracing works
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Perform a full calculation to build the chain
        workbook.CalculateFormula();

        // Select the cell whose precedents and dependents we want to trace
        Cell targetCell = workbook.Worksheets[0].Cells["A2"];

        // ----- Trace precedents (cells that influence the target cell) -----
        IEnumerator precedentsEnum = targetCell.GetPrecedentsInCalculation();
        Console.WriteLine($"Precedents of {targetCell.Name}:");
        if (precedentsEnum != null)
        {
            while (precedentsEnum.MoveNext())
            {
                // Each item is a ReferredArea describing a precedent range
                ReferredArea area = (ReferredArea)precedentsEnum.Current;
                Console.WriteLine(area);
            }
        }
        else
        {
            Console.WriteLine("  None");
        }

        // ----- Trace dependents (cells that depend on the target cell) -----
        // Use recursive = true to include indirect dependents
        IEnumerator dependentsEnum = targetCell.GetDependentsInCalculation(true);
        Console.WriteLine($"Dependents of {targetCell.Name}:");
        if (dependentsEnum != null)
        {
            while (dependentsEnum.MoveNext())
            {
                // Each item is a Cell that depends on the target cell
                Cell dependentCell = (Cell)dependentsEnum.Current;
                Console.WriteLine($"  {dependentCell.Name}");
            }
        }
        else
        {
            Console.WriteLine("  None");
        }

        // Save the workbook (optional, demonstrates lifecycle compliance)
        workbook.Save("output.xlsx");
    }
}