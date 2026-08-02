// Title: Evaluate External Workbook References with Worksheet.CalculateFormula in Aspose.Cells for .NET
// Description: Demonstrates how to create a secondary workbook, assign a value to Sheet1!A2, reference that cell from a primary workbook using the formula =[External.xlsx]Sheet1!$A$2, configure CalculationOptions.LinkedDataSources with the external file, and compute the result both via Worksheet.CalculateFormula and Workbook.CalculateFormula.
// Keywords: Aspose.Cells | Worksheet.CalculateFormula | external reference | LinkedDataSources | CalculationOptions | C# | cross‑workbook formula | .NET | Excel automation | formula evaluation
// Common Searches: Aspose.Cells calculate formula from another file | Worksheet.CalculateFormula external link example | How to use LinkedDataSources in Aspose.Cells | C# evaluate cross‑workbook cell reference | CalculateFormula with external workbook in .NET
// Developer Intent: Compute a cell value that depends on data stored in a separate workbook using Aspose.Cells without opening the source file.
// Use Cases: Generate financial summaries that pull totals from a shared data workbook. | Run batch calculations where the source workbook serves as a static lookup table. | Create reporting tools that consolidate figures from multiple Excel files on the fly.
// AI Prompts: Provide C# code that sets up CalculationOptions.LinkedDataSources to resolve an external cell reference with Worksheet.CalculateFormula. | Show how to handle missing or mismatched external workbook names when evaluating a linked formula in Aspose.Cells. | Explain the difference between direct formula calculation and full workbook recalculation for cross‑file references.

using System;
using Aspose.Cells;

namespace WorksheetCalculateFormulaExternalDemo
{
    // Demonstrates how to create a secondary workbook, assign a value to Sheet1!A2, reference that cell from a primary workbook using the formula =[External.xlsx]Sheet1!$A$2, configure CalculationOptions.LinkedDataSources with the external file, and compute the result both via Worksheet.CalculateFormula and Workbook.CalculateFormula.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create external workbook ----------
                Workbook externalWb = new Workbook();
                // Put a value in Sheet1!A2 (row 1, column 0)
                externalWb.Worksheets[0].Cells["A2"].PutValue(12345);
                // (Optional) give the workbook a name that matches the reference
                externalWb.FileName = "External.xlsx";

                // ---------- Create main workbook ----------
                Workbook mainWb = new Workbook();
                Worksheet sheet = mainWb.Worksheets[0];

                // Set a formula that references the external workbook cell A2
                // Note: the external workbook name must match the name used in the formula
                string externalFormula = "=[External.xlsx]Sheet1!$A$2";
                sheet.Cells["B1"].Formula = externalFormula;

                // ---------- Prepare calculation options ----------
                CalculationOptions calcOptions = new CalculationOptions
                {
                    // Provide the external workbook(s) that formulas may refer to
                    LinkedDataSources = new Workbook[] { externalWb }
                };

                // ---------- Calculate the specific formula directly ----------
                // Demonstrates Worksheet.CalculateFormula(string, CalculationOptions)
                object result = sheet.CalculateFormula(externalFormula, calcOptions);
                Console.WriteLine($"Result of direct calculation: {result}");

                // ---------- Calculate all formulas in the workbook ----------
                // Use Workbook.CalculateFormula(CalculationOptions) to apply linked data sources
                mainWb.CalculateFormula(calcOptions);
                Console.WriteLine($"Result after full workbook calculation (B1): {sheet.Cells["B1"].Value}");

                // Keep console window open
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
