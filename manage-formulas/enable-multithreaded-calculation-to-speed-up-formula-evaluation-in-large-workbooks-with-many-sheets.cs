using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMultiThreadedCalcDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "LargeWorkbook.xlsx";
            const string outputPath = "LargeWorkbook_Calculated.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook.
                Workbook workbook = new Workbook(inputPath);

                // Enable multi‑threaded calculation if the property is available in the current Aspose.Cells version.
                // In newer versions this property exists; otherwise the calculation will run in single‑threaded mode.
                // Uncomment the line below if your Aspose.Cells version supports it.
                // workbook.Settings.EnableMultiThreadedCalculation = true;

                // Perform full calculation of all formulas in the workbook.
                workbook.CalculateFormula();

                // Save the workbook after calculation.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine("Calculation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}