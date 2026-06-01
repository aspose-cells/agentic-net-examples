using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsIterativeCalcDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (must be .xlsx, .xlsm, etc.)
            string sourcePath = "input.xlsx";

            // Path for the output file after enabling iterative calculation
            string outputPath = "output.xlsx";

            // Open the source file as a read‑only stream
            using (FileStream fileStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
            {
                // Create LoadOptions – you can customize options here if needed
                LoadOptions loadOptions = new LoadOptions();

                // Load the workbook from the stream with the specified options
                Workbook workbook = new Workbook(fileStream, loadOptions);

                // Enable iterative calculation to resolve circular references
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

                // Optional: configure iteration limits
                workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
                workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

                // Perform calculation so that formulas are evaluated with the new settings
                workbook.CalculateFormula();

                // Save the modified workbook to a new file
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }

            Console.WriteLine("Workbook loaded, iterative calculation enabled, and saved successfully.");
        }
    }
}