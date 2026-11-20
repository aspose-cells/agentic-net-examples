// Title: Catch CalculationException for circular references while recalculating formulas with Aspose.Cells in C#
// Common Searches: how to catch circular reference errors when using Workbook.CalculateFormula in C# Aspose.Cells | detect circular reference CalculationException during Excel formula recalculation with Aspose.Cells | handle formula recalculation exceptions in Aspose.Cells .NET workbook | Aspose.Cells calculate formula circular reference detection example C# | prevent crash on circular reference while calling CalculateFormula in Aspose.Cells
// Tags: Aspose.Cells Workbook.CalculateFormula error handling | C# calculation exception management with Aspose.Cells | Excel circular reference detection using Aspose.Cells API | save workbook after formula recalculation Aspose.Cells | handling formula errors in .NET Excel automation

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, forces a full formula recalculation using Workbook.CalculateFormula, catches any CalculationException, inspects the message for a circular reference indication, logs appropriate diagnostics, and finally saves the workbook, demonstrating robust handling of circular reference errors in C#.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to load the workbook:");
            Console.WriteLine(ex.Message);
            return;
        }

        try
        {
            // Force recalculation of all formulas in the workbook
            workbook.CalculateFormula();
            Console.WriteLine("Formula recalculation completed successfully.");
        }
        catch (Exception ex)
        {
            // Check if the exception message indicates a circular reference
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("circular reference", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Circular reference detected during calculation:");
                Console.WriteLine(ex.Message);
            }
            else
            {
                // Other calculation errors
                Console.WriteLine("An error occurred during formula calculation:");
                Console.WriteLine(ex.Message);
            }
        }

        try
        {
            // Save the workbook after processing (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to save the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
