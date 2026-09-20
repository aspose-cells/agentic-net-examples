// Title: Detect and handle circular reference errors during Workbook.CalculateFormula in Aspose.Cells for .NET
// AI Prompts: Generate C# code that runs Workbook.CalculateFormula and catches CellsException to log when a circular reference is found. | Show how to abort saving an Excel file in Aspose.Cells if a circular reference exception occurs during formula calculation. | Provide a pattern for inspecting CellsException.Message to differentiate circular reference errors from other calculation failures.
// Common Searches: how to catch circular reference exception when using Aspose.Cells CalculateFormula in C# | Aspose.Cells detect circular reference during workbook recalculation .NET | C# example for handling CellsException caused by circular references in Excel formulas | prevent saving Excel file if Aspose.Cells formula calculation reports circular reference | log circular reference errors while recalculating formulas with Aspose.Cells
// Tags: Workbook.CalculateFormula circular reference handling | CellsException inspection Aspose.Cells | Excel formula recalculation error detection C# | Aspose.Cells circular reference logging | prevent workbook save on calculation failure .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, forces full formula recalculation using Workbook.CalculateFormula, catches CellsException, checks the exception message for a circular reference, logs the issue, and skips saving the workbook when such an error is detected.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
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
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Force full calculation of all formulas in the workbook
            workbook.CalculateFormula();
            Console.WriteLine("Formula recalculation completed successfully.");
        }
        catch (CellsException ex)
        {
            // Detect circular reference errors by inspecting the exception message
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("circular reference", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Circular reference detected during formula recalculation.");
                // Additional handling logic can be placed here (e.g., logging, user notification)
            }
            else
            {
                // Re-throw if it's a different calculation exception
                throw;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred during calculation: {ex.Message}");
            return;
        }

        try
        {
            // Save the workbook after successful calculation (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
