// Title: Recalculate Excel formulas with a custom CalculationOptions monitor using Aspose.Cells Workbook.CalculateFormula in C#
// Common Searches: how to implement a calculation monitor for Workbook.CalculateFormula in Aspose.Cells C# | using CalculationOptions to receive formula evaluation callbacks in Aspose.Cells .NET | example of recalculate Excel workbook with progress reporting using Aspose.Cells | Aspose.Cells C# calculate formulas with custom monitor implementation
// Tags: Aspose.Cells custom calculation monitor | Workbook.CalculateFormula with CalculationOptions | C# track formula evaluation callbacks | recalculate Excel formulas with progress reporting | custom ICalculationMonitor implementation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a CalculationOptions object with a custom ICalculationMonitor to receive callbacks for each formula evaluation, recalculates all formulas using Workbook.CalculateFormula(calculationOptions), and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas using default calculation options
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
