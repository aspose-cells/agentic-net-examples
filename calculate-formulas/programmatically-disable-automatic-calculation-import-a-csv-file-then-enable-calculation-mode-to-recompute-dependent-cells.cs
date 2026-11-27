// Title: Disable automatic calculation, import a CSV file, then enable calculation and recalculate formulas with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load CSV without triggering formula evaluation | Set manual calculation mode before importing CSV using Aspose.Cells | Enable automatic calculation after CSV import and recalculate all formulas Aspose.Cells | How to use LoadOptions with LoadFormat.Csv to import data into a workbook in Aspose.Cells | Recalculate dependent cells after loading CSV data with Aspose.Cells .NET
// Tags: set CalcModeType.Manual Aspose.Cells | load CSV with LoadFormat.Csv Aspose.Cells | switch to CalcModeType.Automatic Aspose.Cells | Workbook.CalculateFormula method C# | save workbook as XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImportDemo
{
    // The example creates a Workbook, sets its FormulaSettings.CalculationMode to Manual to prevent formula evaluation, loads a CSV file via a FileStream using LoadOptions with LoadFormat.Csv, switches the calculation mode back to Automatic, calls Workbook.CalculateFormula to recompute dependent formulas, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: AsposeCellsCsvImportDemo <inputCsvPath> [outputExcelPath]");
                return;
            }

            string csvPath = args[0];
            string outputPath = args.Length >= 2 ? args[1] : "Result.xlsx";

            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
                return;
            }

            try
            {
                // 1. Create a new Workbook instance with manual calculation mode.
                //    This prevents any formulas from being evaluated while we load the CSV.
                Workbook workbook = new Workbook();
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;
                Console.WriteLine("Automatic calculation disabled (Manual mode).");

                // 2. Load the CSV file into the workbook.
                //    Use LoadOptions to specify that the source format is CSV.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
                // The constructor that accepts a stream and LoadOptions is used to avoid
                // any implicit calculation that might happen when loading from a file path.
                using (FileStream csvStream = new FileStream(csvPath, FileMode.Open, FileAccess.Read))
                {
                    // Load the CSV data into the workbook.
                    workbook = new Workbook(csvStream, loadOptions);
                }
                Console.WriteLine($"CSV file '{csvPath}' imported successfully.");

                // 3. Switch calculation mode back to Automatic so that dependent formulas are recomputed.
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
                Console.WriteLine("Automatic calculation re-enabled.");

                // 4. Recalculate all formulas in the workbook.
                workbook.CalculateFormula();
                Console.WriteLine("Formulas recalculated.");

                // 5. Save the resulting workbook to the desired output format (XLSX by default).
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
