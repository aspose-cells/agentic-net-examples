// Title: Aspose.Cells for .NET – Suppress External Link Errors When Calculating Formulas
// Description: Demonstrates how to configure a Workbook to ignore missing external‑link errors by setting CalculationOptions.IgnoreError to true, applying the option to workbook.CalculateFormula, and saving the result. Includes correct external reference syntax and error‑free execution.
// Keywords: Aspose.Cells ignore external link errors | CalculationOptions IgnoreError .NET | suppress formula errors Aspose.Cells | missing external workbook reference | disable external link errors Excel | Aspose.Cells calculate formula without links
// Common Searches: Aspose.Cells ignore external link errors | CalculationOptions.IgnoreError example C# | prevent formula failure missing external workbook Aspose.Cells | how to suppress external reference errors in Excel using Aspose.Cells | calculate formulas without external links Aspose.Cells .NET
// Developer Intent: Avoid calculation exceptions caused by broken or unavailable external workbook references.
// Use Cases: Generate reports that contain external formulas when the source files are not deployed. | Process user‑uploaded Excel templates with broken links on a server without raising errors. | Run batch calculations on workbooks in a cloud service where external links cannot be resolved.
// AI Prompts: Show code that sets CalculationOptions.IgnoreError and logs cells that had errors suppressed. | Explain how to combine IgnoreError with Recursive and EnableIterativeCalculation options in Aspose.Cells. | Provide a step‑by‑step guide to disable external link updating while still allowing normal formula evaluation.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to configure a Workbook to ignore missing external‑link errors by setting CalculationOptions.IgnoreError to true, applying the option to workbook.CalculateFormula, and saving the result. Includes correct external reference syntax and error‑free execution.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a formula that references an external workbook (which may be missing)
            Worksheet sheet = workbook.Worksheets[0];
            // Correct external reference syntax: ='[FileName]SheetName'!CellReference
            sheet.Cells["A1"].Formula = "='[NonExistingFile.xlsx]Sheet1'!$A$1";

            // Set calculation options to ignore errors (including external link errors)
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true // suppress errors during formula evaluation
            };

            // Calculate all formulas using the configured options
            workbook.CalculateFormula(calcOptions);

            // Display the result of the cell after calculation
            Console.WriteLine("A1 value after calculation: " + sheet.Cells["A1"].StringValue);

            // Save the workbook (optional)
            string outputPath = "IgnoreExternalLinkErrors.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
