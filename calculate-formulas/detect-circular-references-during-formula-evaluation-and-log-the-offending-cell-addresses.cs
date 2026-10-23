// Title: Detect Circular References and Log Affected Cell Addresses When Calculating Formulas with Aspose.Cells for .NET
// Common Searches: how to catch circular reference errors while using Workbook.CalculateFormula in C# | log the cell addresses that cause circular reference exceptions in Aspose.Cells | detect circular references in Excel formulas with Aspose.Cells .NET API | handle formula calculation errors and retrieve offending cells using Aspose.Cells | Aspose.Cells calculate formula circular reference detection example C#
// Tags: Workbook.CalculateFormula exception handling | Aspose.Cells circular reference detection | log cell addresses for formula errors C# | Excel formula evaluation error logging Aspose | handle circular reference errors Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, runs Workbook.CalculateFormula to evaluate all formulas, catches any exceptions thrown due to circular references, writes the exception message (including the offending cell addresses) to the console, and then saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to load workbook:");
            Console.WriteLine(ex.Message);
            return;
        }

        try
        {
            // Evaluate all formulas in the workbook.
            // If a circular reference exists, Aspose.Cells will throw an exception.
            workbook.CalculateFormula();
        }
        catch (Exception ex)
        {
            // Handle any calculation errors, including circular references
            Console.WriteLine("Error during formula calculation:");
            Console.WriteLine(ex.Message);
        }

        try
        {
            // Save the workbook after evaluation (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to save workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
