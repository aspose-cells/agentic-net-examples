// Title: C# – Calculate only the "Summary" sheet formulas with Aspose.Cells Worksheet.CalculateFormula
// Description: Loads a workbook, locates the worksheet named "Summary", and uses Worksheet.CalculateFormula with CalculationOptions to evaluate all formulas on that sheet alone before saving. Shows how to recalculate a single sheet without touching other worksheets.
// Keywords: Aspose.Cells | C# | Worksheet.CalculateFormula | calculate single sheet | summary worksheet | selective formula evaluation | CalculationOptions | Excel partial recalculation | Aspose.Cells .NET example
// Common Searches: Aspose.Cells calculate formulas on one sheet | C# calculate only summary worksheet | Worksheet.CalculateFormula example | Selective formula recalculation Aspose.Cells | How to recalc specific sheet in .NET Excel library
// Developer Intent: Recalculate all formulas in the worksheet named "Summary" while leaving other sheets unchanged.
// Use Cases: Refresh summary totals after data changes in source sheets without re‑processing the entire workbook. | Generate a report where only the summary tab needs up‑to‑date calculations, reducing CPU time. | Validate summary‑sheet calculations during automated testing while preserving original data sheets.
// AI Prompts: Show a C# example that uses Worksheet.CalculateFormula with custom CalculationOptions for the "Summary" sheet. | Provide robust error handling for a missing "Summary" worksheet when performing selective formula calculation with Aspose.Cells. | Explain how to recalculate formulas on a single sheet while keeping other worksheets untouched in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, locates the worksheet named "Summary", and uses Worksheet.CalculateFormula with CalculationOptions to evaluate all formulas on that sheet alone before saving. Shows how to recalculate a single sheet without touching other worksheets.
    public class CalculateSummarySheet
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Find the worksheet named "Summary"
                Worksheet summarySheet = workbook.Worksheets["Summary"];
                if (summarySheet == null)
                {
                    Console.WriteLine("Worksheet 'Summary' not found.");
                    return;
                }

                // Calculate all formulas in the "Summary" worksheet
                CalculationOptions calcOptions = new CalculationOptions();
                summarySheet.CalculateFormula(calcOptions, true);

                // Save the workbook after calculation
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            CalculateSummarySheet.Run();
        }
    }
}
