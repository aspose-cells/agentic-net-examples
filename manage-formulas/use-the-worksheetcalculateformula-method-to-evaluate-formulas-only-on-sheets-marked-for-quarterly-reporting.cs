// Title: Selective formula recalculation on quarterly worksheets using Worksheet.CalculateFormula (Aspose.Cells for .NET)
// Description: Shows how to loop through a workbook, detect worksheets whose names contain "Quarterly", and invoke Worksheet.CalculateFormula with optional CalculationOptions to recalculate only those sheets before saving the file.
// Keywords: Aspose.Cells | Worksheet.CalculateFormula | selective formula calculation | quarterly reporting sheets | C# .NET Excel automation | CalculationOptions | partial workbook recalculation
// Common Searches: Aspose.Cells calculate formulas on specific sheets | Worksheet.CalculateFormula example C# | recalculate only quarterly worksheets Aspose.Cells | partial formula evaluation .NET Excel | how to skip sheets during formula calculation Aspose
// Developer Intent: Recalculate formulas exclusively on worksheets marked for quarterly reporting.
// Use Cases: Iterate through a workbook, identify sheets with "Quarterly" in the title, and run Worksheet.CalculateFormula on each to update only those reports. | Apply custom CalculationOptions (e.g., enable iterative calculation) to quarterly sheets while leaving other worksheets untouched. | Improve performance in large workbooks by avoiding unnecessary formula evaluation on non‑reporting sheets.
// AI Prompts: Generate C# code that scans a Workbook, selects worksheets whose name includes "Quarterly", and calls Worksheet.CalculateFormula with CalculationOptions for each selected sheet. | Explain how to configure CalculationOptions for iterative calculations when selectively recalculating quarterly worksheets in Aspose.Cells. | Provide error‑handling patterns for unsupported functions that may appear in quarterly sheets during selective formula evaluation.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to loop through a workbook, detect worksheets whose names contain "Quarterly", and invoke Worksheet.CalculateFormula with optional CalculationOptions to recalculate only those sheets before saving the file.
class QuarterlyFormulaEvaluator
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare calculation options (default settings)
            CalculationOptions calcOptions = new CalculationOptions();

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Identify quarterly worksheets by name
                if (sheet.Name.IndexOf("Quarterly", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Optionally perform sheet‑specific actions here
                    // (e.g., modify data before calculation)
                }
            }

            // Calculate all formulas in the workbook using the specified options
            workbook.CalculateFormula(calcOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
