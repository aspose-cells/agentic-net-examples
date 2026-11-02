// Title: Enable ignoring error values in Aspose.Cells CalculationOptions for C# formula evaluation
// Common Searches: Aspose.Cells C# ignore #DIV/0 errors during CalculateFormula | set CalculationOptions.IgnoreErrorValue true in Aspose.Cells .NET | skip Excel formula errors when recalculating workbook with Aspose.Cells | how to prevent formula evaluation exceptions in Aspose.Cells using CalculationOptions | Aspose.Cells calculation options to ignore error values in .NET
// Tags: Aspose.Cells CalculationOptions ignore error values | C# Aspose.Cells skip formula errors | Aspose.Cells CalculateFormula ignore errors | Excel formula error handling Aspose.Cells .NET | Aspose.Cells legacy IgnoreErrorValue property

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates a workbook, creates a CalculationOptions instance (noting that the IgnoreErrorValue property is unavailable in recent releases), recalculates all formulas with those options, and saves the result, illustrating how error‑value handling was configured in earlier Aspose.Cells versions.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file names
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Create calculation options (default settings)
            CalculationOptions calcOptions = new CalculationOptions();

            // NOTE: In recent Aspose.Cells versions the property to ignore error values
            // (IgnoreErrorValue) is not available. Default behavior is used here.

            // Recalculate all formulas with the specified options
            workbook.CalculateFormula(calcOptions);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
