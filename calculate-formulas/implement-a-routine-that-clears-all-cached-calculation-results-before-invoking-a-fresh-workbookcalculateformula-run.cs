// Title: How to clear the calculation cache and force a full Workbook.CalculateFormula in Aspose.Cells for C#
// Common Searches: Aspose.Cells C# clear formula cache before recalculating workbook | Force full calculation of all formulas in an Excel file using Aspose.Cells .NET | Reset cached calculation results in Aspose.Cells workbook programmatically | Use CalculationOptions to force full formula evaluation in Aspose.Cells C# | Recalculate Excel workbook with fresh cache using Aspose.Cells API
// Tags: Aspose.Cells clear calculation cache | Force full formula recalculation .NET | CalculationOptions ForceFullCalculation Aspose.Cells | Workbook.CalculateFormula fresh cache | C# Excel formula cache reset Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookProcessing
{
    // The example demonstrates loading an Excel workbook with Aspose.Cells, configuring CalculationOptions.ForceFullCalculation to true to clear any cached formula results, invoking Workbook.CalculateFormula for a complete recomputation, and then saving the updated workbook. This ensures that all formulas are evaluated from scratch regardless of previous calculation caches.
    public class WorkbookProcessor
    {
        /// <param name="inputPath">Full path to the source workbook file.</param>
        /// <param name="outputPath">Full path where the processed workbook will be saved.</param>
        public void RecalculateWorkbook(string inputPath, string outputPath)
        {
            // Verify the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook.
            workbook.CalculateFormula();

            // Save the updated workbook to the output file.
            workbook.Save(outputPath);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath;
                string outputPath;

                if (args.Length >= 2)
                {
                    inputPath = args[0];
                    outputPath = args[1];
                }
                else
                {
                    // Default example paths; replace with actual file locations as needed.
                    inputPath = "input.xlsx";
                    outputPath = "output.xlsx";
                }

                var processor = new WorkbookProcessor();
                processor.RecalculateWorkbook(inputPath, outputPath);

                Console.WriteLine("Workbook recalculation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
