// Title: C# – Calculate formulas only on the 'Summary' worksheet using Aspose.Cells
// Description: Loads an Excel file, verifies the presence of a worksheet named "Summary", and runs Worksheet.CalculateFormula with includeDependentCells set to true so that only formulas on that sheet (and any cells they depend on in other sheets) are evaluated before saving the workbook.
// Keywords: Aspose.Cells | Worksheet.CalculateFormula | partial workbook recalculation | C# Excel formula calculation | .NET Excel performance | summary sheet calculation | dependent cells across sheets
// Common Searches: Aspose.Cells calculate formulas on a single sheet | C# calculate only Summary worksheet formulas | Worksheet.CalculateFormula include dependent cells | partial recalculation Excel using Aspose.Cells | how to recalc specific sheet Aspose.Cells .NET
// Developer Intent: Recalculate formulas exclusively on the 'Summary' sheet while optionally updating dependent cells in other worksheets.
// Use Cases: Refresh totals on a dashboard sheet after data changes elsewhere without reprocessing the entire workbook. | Generate a financial summary report where only the summary tab needs updated calculations before export. | Speed up automated batch jobs on large workbooks by limiting formula evaluation to a single worksheet.
// AI Prompts: Provide C# code that uses Aspose.Cells to calculate formulas on a worksheet named "Summary" and includes dependent cells from other sheets. | Explain how to handle a missing "Summary" worksheet when calling Worksheet.CalculateFormula in Aspose.Cells. | Show how to configure CalculationOptions for partial formula evaluation with Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file, verifies the presence of a worksheet named "Summary", and runs Worksheet.CalculateFormula with includeDependentCells set to true so that only formulas on that sheet (and any cells they depend on in other sheets) are evaluated before saving the workbook.
    public class CalculateSummarySheet
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Locate the "Summary" worksheet
                Worksheet summarySheet = workbook.Worksheets["Summary"];
                if (summarySheet == null)
                {
                    Console.WriteLine("Worksheet 'Summary' not found.");
                    return;
                }

                // Set calculation options (default is sufficient)
                CalculationOptions calcOptions = new CalculationOptions();

                // Calculate formulas in the "Summary" sheet, including dependent cells in other sheets
                summarySheet.CalculateFormula(calcOptions, true);

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Formulas in 'Summary' worksheet have been calculated and workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
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
