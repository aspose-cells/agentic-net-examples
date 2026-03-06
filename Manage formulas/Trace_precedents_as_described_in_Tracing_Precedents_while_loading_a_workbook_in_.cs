using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TracePrecedentsDemo
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook with default options (formula parsing on open)
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = true; // ensure formulas are parsed during load
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Enable calculation chain to allow tracing of calculation precedents
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Perform full calculation so that precedents in calculation are available
            workbook.CalculateFormula();

            // Choose the cell whose precedents we want to trace (e.g., A2)
            Cell targetCell = workbook.Worksheets[0].Cells["A2"];

            // Get the calculation precedents (only those that affect the result)
            IEnumerator precedentsEnum = targetCell.GetPrecedentsInCalculation();

            Console.WriteLine($"Precedents influencing the calculation of cell {targetCell.Name}:");

            if (precedentsEnum != null)
            {
                while (precedentsEnum.MoveNext())
                {
                    // Each item is a ReferredArea describing a referenced range or cell
                    ReferredArea area = (ReferredArea)precedentsEnum.Current;

                    // Build a readable representation of the area
                    string areaDesc = "";

                    if (area.IsExternalLink)
                    {
                        areaDesc += $"[{area.ExternalFileName}]";
                    }

                    areaDesc += $"{area.SheetName}!";

                    // Start cell
                    areaDesc += CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);

                    // If the area spans multiple cells, add the end cell
                    if (area.IsArea)
                    {
                        areaDesc += $":{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}";
                    }

                    Console.WriteLine(areaDesc);
                }
            }
            else
            {
                Console.WriteLine("No precedents found (cell may not contain a formula).");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TracePrecedentsDemo.Run();
        }
    }
}