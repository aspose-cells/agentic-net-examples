// Title: Detect circular references in Excel formulas using Aspose.Cells for .NET during workbook.CalculateFormula
// Common Searches: how to catch circular reference errors when calculating Excel formulas with Aspose.Cells C# | Aspose.Cells workbook.CalculateFormula throws exception for circular references | C# detect circular reference in Excel workbook using Aspose.Cells | log circular reference detection message with Aspose.Cells .NET | handle formula calculation errors in Aspose.Cells C# example
// Tags: Aspose.Cells circular reference detection | C# Excel formula calculation exception handling | Aspose.Cells workbook.CalculateFormula error handling | log circular reference detection Aspose.Cells | detect circular reference in Excel using .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, forces formula calculation with Aspose.Cells, catches any exception that indicates a circular reference, outputs a detection message, and then saves the workbook.
class CircularReferenceDetector
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook
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
            // Force calculation of all formulas in the workbook.
            // If a circular reference exists, Aspose.Cells throws an exception.
            workbook.CalculateFormula();
            Console.WriteLine("No circular references were found.");
        }
        catch (Exception ex)
        {
            // Any exception during calculation is treated as a circular reference detection.
            Console.WriteLine("Circular reference detected or calculation error:");
            Console.WriteLine(ex.Message);
        }

        try
        {
            // Save the workbook (if needed)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to save the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
