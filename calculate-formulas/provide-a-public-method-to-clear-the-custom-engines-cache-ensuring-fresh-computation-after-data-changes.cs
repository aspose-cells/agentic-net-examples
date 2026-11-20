// Title: Clear the calculation cache in Aspose.Cells with C# to force fresh formula evaluation
// Common Searches: C# Aspose.Cells how to reset the formula calculation cache | programmatically clear cached formula results in a workbook using Aspose.Cells | force all formulas to recalculate after data changes with Aspose.Cells .NET | Aspose.Cells clear calculation cache before saving workbook | recalculate formulas without opening Excel using Aspose.Cells C#
// Tags: Aspose.Cells calculation cache reset | Workbook.CalculateFormula usage | clear formula cache C# | force formula recomputation Aspose.Cells | clear cache before workbook save

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example defines a SpreadsheetHelper class with a ClearCalculationCache method that calls Workbook.CalculateFormula inside a try‑catch block, forcing all formulas to be recomputed and clearing any cached results. The program loads or creates a workbook, invokes this method, and saves the file.
    public class SpreadsheetHelper
    {
        private readonly Workbook _workbook;

        public SpreadsheetHelper(Workbook workbook)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        public void ClearCalculationCache()
        {
            try
            {
                // Recalculate all formulas to clear any cached results.
                _workbook.CalculateFormula();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error recalculating formulas: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            const string inputPath = "sample.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                Workbook workbook;

                // Load existing workbook if present; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                var helper = new SpreadsheetHelper(workbook);
                helper.ClearCalculationCache();

                // Save the workbook to verify the process completed.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
