using System;
using Aspose.Cells;

namespace AsposeCellsFormulaModeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook (XLSX format) using default load options
            Workbook workbook = new Workbook(inputPath);

            // Set the desired formula calculation mode.
            // Options: Automatic, AutomaticExceptTable, Manual
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Optional: display the mode that was set
            Console.WriteLine("Calculation Mode set to: " + workbook.Settings.FormulaSettings.CalculationMode);

            // Save the workbook back to a new file (or overwrite the original)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with updated calculation mode.");
        }
    }
}