// Title: Ignore Missing External Links During Formula Calculation with Aspose.Cells for .NET
// Description: Shows how to enable CalculationOptions.IgnoreError in Aspose.Cells so formulas that reference external workbooks are evaluated without triggering link‑related exceptions, then saves the result.
// Keywords: Aspose.Cells ignore external links | CalculationOptions IgnoreError .NET | skip missing external references Excel | Aspose.Cells calculate formula without external workbook | C# Aspose.Cells external link error handling
// Common Searches: Aspose.Cells ignore external link errors | CalculationOptions.IgnoreError example C# | prevent formula failure due to missing external workbook Aspose.Cells | how to calculate formulas with broken external links in .NET
// Developer Intent: Suppress errors caused by absent external references while calculating formulas in an Aspose.Cells workbook.
// Use Cases: Run batch calculations on templates that contain placeholder external links. | Generate reports from workbooks where external sources are unavailable or intentionally omitted. | Process large collections of files with broken links without interrupting the calculation workflow.
// AI Prompts: Provide a C# snippet that sets CalculationOptions.IgnoreError to true, calculates all formulas, and saves the workbook. | Explain the impact of enabling IgnoreError on external link handling versus other calculation errors in Aspose.Cells. | Show how to catch and log only non‑ignored calculation errors while skipping missing external link exceptions.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to enable CalculationOptions.IgnoreError in Aspose.Cells so formulas that reference external workbooks are evaluated without triggering link‑related exceptions, then saves the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add a formula that references an external workbook
            Worksheet sheet = workbook.Worksheets[0];
            // Correct external reference syntax: ='[ExternalSource.xlsx]Sheet1'!A1
            sheet.Cells["A1"].Formula = "='[ExternalSource.xlsx]Sheet1'!A1";

            // Configure calculation options to ignore errors (including missing external links)
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true
            };

            // Perform formula calculation with the configured options
            workbook.CalculateFormula(calcOptions);

            // Save the workbook after calculation
            string outputPath = "Output_IgnoringExternalLinkErrors.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
