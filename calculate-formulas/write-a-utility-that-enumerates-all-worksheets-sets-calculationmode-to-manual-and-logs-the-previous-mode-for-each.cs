using System;
using Aspose.Cells;

namespace AsposeCellsUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Store the original calculation mode (it's a workbook‑level setting)
            CalcModeType originalMode = workbook.Settings.FormulaSettings.CalculationMode;

            // Enumerate all worksheets and log the previous mode for each
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\": Previous CalculationMode = {originalMode}");
            }

            // Set the calculation mode to Manual for the entire workbook
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}