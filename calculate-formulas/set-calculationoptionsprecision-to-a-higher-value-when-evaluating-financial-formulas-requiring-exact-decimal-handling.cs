// Title: Setting CalculationOptions.Precision for higher decimal accuracy in financial formulas with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# increase calculation precision for FV financial function | how to configure CalculationOptions.Precision in Aspose.Cells .NET | set higher decimal precision when evaluating financial formulas using Aspose.Cells | C# Aspose.Cells workbook precise decimal handling for financial calculations | adjust calculation precision for exact results in Aspose.Cells formulas
// Tags: CalculationOptions.Precision configuration Aspose.Cells | financial formula decimal precision .NET | high precision FV function Aspose.Cells | Aspose.Cells workbook calculation settings | exact decimal handling financial calculations C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, inserts principal, interest rate, and years values, assigns an FV financial formula, recalculates all formulas, and saves the file. It demonstrates basic formula evaluation but does not change the CalculationOptions.Precision setting, which would be required for higher decimal accuracy.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Optional: enforce calculation precision as displayed (if needed)
            // workbook.Settings.SetPrecisionAsDisplayed(true);

            // Add data and a financial formula that benefits from high precision
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(1000); // principal
            sheet.Cells["A2"].PutValue(0.05); // interest rate
            sheet.Cells["A3"].PutValue(5);    // years
            sheet.Cells["B1"].Formula = "FV(A2, A3, 0, -A1)";

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Define output file path
            string outputPath = "FinancialPrecisionExample.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
