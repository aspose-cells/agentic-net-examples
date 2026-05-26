using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook file
            string inputPath = "input.xlsx";

            // Load the workbook from disk using the string constructor (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Set the calculation mode to Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Save the workbook back to disk (save rule)
            string outputPath = "output_manual.xlsx";
            workbook.Save(outputPath);
        }
    }
}