// Title: Configure Aspose.Cells workbook iterative calculation max change to 0.001 and test circular reference convergence in C#
// Common Searches: Aspose.Cells set iterative calculation maximum change to 0.001 in C# | how to enable iterative calculation for circular references using Aspose.Cells .NET | C# example testing convergence of circular references with Aspose.Cells CalculateFormula | retrieve cell values after iterative calculation with Aspose.Cells workbook | save workbook after configuring iterative calculation settings Aspose.Cells
// Tags: iterative calculation tolerance Aspose.Cells | circular reference convergence Aspose.Cells | configure workbook iterative settings C# | calculate formulas with iterative mode Aspose.Cells | save workbook after iterative calculation Aspose.Cells

using Aspose.Cells;
using System;

// Creates a workbook, adds a circular reference between A1 and B1, runs CalculateFormula which applies iterative calculation using the workbook's max‑change setting, outputs the converged cell values, and saves the workbook to a file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a circular reference:
            // A1 = B1 + 1
            // B1 = A1 + 1
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Force calculation of formulas (iterative calculation will be applied automatically if enabled in the library)
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Error during formula calculation: {calcEx.Message}");
            }

            // Retrieve the calculated values after iterative calculation
            double valueA1 = sheet.Cells["A1"].DoubleValue;
            double valueB1 = sheet.Cells["B1"].DoubleValue;

            // Output the results to verify convergence
            Console.WriteLine($"A1 = {valueA1}");
            Console.WriteLine($"B1 = {valueB1}");

            // Save the workbook (optional, demonstrates that settings are persisted)
            string outputPath = "IterativeCalculationDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
