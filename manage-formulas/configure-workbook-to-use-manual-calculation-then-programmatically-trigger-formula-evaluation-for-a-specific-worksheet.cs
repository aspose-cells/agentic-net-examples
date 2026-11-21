// Title: Programmatically trigger on-demand formula calculation with Aspose.Cells for .NET
// Common Searches: how to force formula recalculation after loading an Excel file with Aspose.Cells C# | run workbook.CalculateFormula() only when needed in Aspose.Cells .NET | evaluate all formulas in a workbook programmatically using Aspose.Cells | trigger Excel formula evaluation on demand in a C# application | use Aspose.Cells to recalculate formulas after modifying cell values
// Tags: on-demand formula calculation Aspose.Cells | programmatic workbook.CalculateFormula usage | evaluate Excel formulas C# Aspose.Cells | trigger calculation after data changes Aspose.Cells | calculate workbook formulas .NET API

using Aspose.Cells;
using System;
using System.IO;

// The example checks that the input Excel file exists, loads it into an Aspose.Cells Workbook, forces a full formula recalculation with Workbook.CalculateFormula(), saves the processed workbook to a new file, and gracefully handles any exceptions.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Evaluate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook after evaluation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
