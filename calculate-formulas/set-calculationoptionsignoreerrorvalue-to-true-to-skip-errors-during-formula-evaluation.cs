// Title: Setting CalculationOptions.IgnoreErrorValue = true in Aspose.Cells C# to skip formula errors
// Common Searches: Aspose.Cells C# ignore #DIV/0 errors during CalculateFormula | set CalculationOptions.IgnoreErrorValue to true in .NET workbook calculation | skip Excel error values when evaluating formulas with Aspose.Cells | configure Aspose.Cells calculation options to suppress formula errors in C# | how to prevent formula error propagation in Aspose.Cells .NET
// Tags: Aspose.Cells CalculationOptions IgnoreErrorValue true | skip formula errors Aspose.Cells C# | suppress Excel error values during calculation .NET | configure workbook calculation options Aspose.Cells | disable error propagation in formula evaluation C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates an in‑memory workbook, sets CalculationOptions.IgnoreErrorValue to true to ignore error values such as #DIV/0 during formula evaluation, forces calculation with CalculateFormula, and then saves the workbook to output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // (Optional) Force calculation to apply default settings
            workbook.CalculateFormula();

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.xlsx");

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
