// Title: Detect and Handle Circular Reference Errors When Recalculating Formulas with Aspose.Cells in C#
// Common Searches: how to catch circular reference errors while using Workbook.CalculateFormula in C# | Aspose.Cells detect circular reference during formula recalculation | C# code to handle CalculationException for circular references in an Excel workbook | prevent crash when recalculating formulas that contain circular references using Aspose.Cells | log circular reference message from CalculateFormula exception Aspose.Cells .NET
// Tags: Aspose.Cells circular reference detection | Workbook.CalculateFormula exception handling | C# Excel formula recalculation error handling | detect circular reference Aspose.Cells .NET | handle CalculationException Aspose.Cells | Excel circular reference handling C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, attempts to recalculate all formulas with Workbook.CalculateFormula, catches CalculationException, checks the message for a circular reference indication, and provides custom handling while safely saving the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Recalculate all formulas in the workbook
        try
        {
            workbook.CalculateFormula();
            Console.WriteLine("Formula recalculation completed successfully.");
        }
        catch (Exception ex)
        {
            // Detect circular reference errors by inspecting the message
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("circular reference", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Circular reference detected during formula recalculation.");
                // Custom handling for circular references can be placed here
            }
            else
            {
                Console.WriteLine($"Error during formula recalculation: {ex.Message}");
            }
        }

        // Ensure the output directory exists
        try
        {
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to prepare output directory: {ex.Message}");
            return;
        }

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
