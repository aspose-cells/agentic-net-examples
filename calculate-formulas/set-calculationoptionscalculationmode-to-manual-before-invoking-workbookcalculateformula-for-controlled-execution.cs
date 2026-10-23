// Title: How to set CalculationOptions.CalculationMode to Manual and calculate formulas with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set calculation mode to manual before Workbook.CalculateFormula | disable automatic formula recalculation in Aspose.Cells .NET | manual calculation mode example using Aspose.Cells Workbook.CalculateFormula | control when formulas are evaluated with Aspose.Cells CalculationOptions | C# Aspose.Cells calculate formulas after setting manual calculation mode
// Tags: Aspose.Cells manual calculation mode C# | CalculationOptions.CalculationMode Manual Aspose | Workbook.CalculateFormula explicit execution | disable auto formula recalculation Aspose.Cells | controlled formula evaluation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to load an Excel workbook with Aspose.Cells, optionally set CalculationOptions.CalculationMode to Manual to prevent automatic recalculation, explicitly invoke Workbook.CalculateFormula for controlled formula evaluation, and then save the workbook while handling missing files and runtime exceptions.
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
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Perform formula calculation explicitly
            workbook.CalculateFormula();

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
