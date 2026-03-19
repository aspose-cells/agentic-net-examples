using System;
using System.Collections;
using Aspose.Cells;

class TracingPrecedentsDemo
{
    static void Main()
    {
        // Load the XLSX workbook
        string inputPath = "input.xlsx";
        LoadOptions loadOptions = new LoadOptions();
        // Parse formulas when opening the file (default is true)
        loadOptions.ParsingFormulaOnOpen = true;
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Enable calculation chain to allow tracing of precedents
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Perform full calculation so the chain is built
        workbook.CalculateFormula();

        // Select the cell whose precedents we want to trace (example: A2)
        Cell targetCell = workbook.Worksheets[0].Cells["A2"];

        // Retrieve precedents that actually influence the calculation
        IEnumerator precedents = targetCell.GetPrecedentsInCalculation();

        if (precedents != null)
        {
            Console.WriteLine($"Precedents of cell {targetCell.Name}:");
            while (precedents.MoveNext())
            {
                ReferredArea area = (ReferredArea)precedents.Current;
                Console.WriteLine(area);
            }
        }
        else
        {
            Console.WriteLine($"Cell {targetCell.Name} has no calculation precedents.");
        }

        // Optionally save the workbook after processing
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}