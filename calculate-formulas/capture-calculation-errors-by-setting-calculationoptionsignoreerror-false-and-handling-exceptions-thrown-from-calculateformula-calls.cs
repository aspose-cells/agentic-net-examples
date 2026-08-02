using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CaptureCalculationErrors
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set a formula that will generate an error (non‑existent function)
                sheet.Cells["A1"].Formula = "=NONEXISTENTFUNC(1)";

                // Configure calculation options to NOT ignore errors
                CalculationOptions options = new CalculationOptions
                {
                    IgnoreError = false
                };

                // Perform calculation and capture any exception thrown
                try
                {
                    workbook.CalculateFormula(options);
                    Console.WriteLine("Calculation completed without errors.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Calculation error captured: " + ex.Message);
                }

                // Re‑calculate while ignoring errors to allow further processing
                options.IgnoreError = true;
                workbook.CalculateFormula(options);
                Console.WriteLine("Recalculated with errors ignored.");

                // Save the workbook (optional)
                string outputPath = "CaptureCalculationErrors.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unexpected error: {e.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}