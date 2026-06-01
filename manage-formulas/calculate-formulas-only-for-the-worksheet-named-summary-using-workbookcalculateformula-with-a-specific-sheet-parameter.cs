using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CalculateSummarySheet
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheet named "Summary"
                Worksheet summarySheet = workbook.Worksheets["Summary"];
                if (summarySheet == null)
                {
                    Console.WriteLine("Worksheet 'Summary' not found.");
                    return;
                }

                // Create calculation options (default options are sufficient)
                CalculationOptions calcOptions = new CalculationOptions();

                // Calculate formulas only on the "Summary" sheet
                summarySheet.CalculateFormula(calcOptions, false);

                // Save the workbook after calculation
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}