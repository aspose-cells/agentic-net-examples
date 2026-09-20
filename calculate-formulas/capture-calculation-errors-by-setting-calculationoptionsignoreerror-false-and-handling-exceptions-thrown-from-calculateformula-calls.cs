// Title: How to capture formula calculation errors in Aspose.Cells for .NET by disabling IgnoreError and handling exceptions
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, sets CalculationOptions.IgnoreError to false, runs Workbook.CalculateFormula, and logs any thrown exceptions. | Show an example of handling load, calculation, and save failures when processing a workbook using Aspose.Cells calculation options that do not ignore errors.
// Common Searches: Aspose.Cells calculate formula without ignoring errors and capture exception details | C# Aspose.Cells CalculationOptions.IgnoreError false example | How to log formula evaluation errors when using Workbook.CalculateFormula in .NET | Detect and handle Excel formula errors with Aspose.Cells calculation engine
// Tags: Aspose.Cells calculation options ignoreerror false | C# Aspose.Cells formula error handling | Workbook.CalculateFormula exception capture | Excel workbook load and save error handling with Aspose.Cells | disable ignoreerror in Aspose.Cells calculation

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file using Aspose.Cells, configures CalculationOptions to disable error ignoring, executes Workbook.CalculateFormula inside a try‑catch block to capture any formula evaluation errors, and then saves the workbook while reporting load, calculation, and save failures.
class CalculationErrorCapture
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
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
            Console.WriteLine("Failed to load workbook:");
            Console.WriteLine(ex.Message);
            return;
        }

        // Configure calculation options (do NOT ignore errors)
        var calcOptions = new CalculationOptions
        {
            IgnoreError = false
        };

        try
        {
            // Perform formula calculation with the specified options
            workbook.CalculateFormula(calcOptions);
            Console.WriteLine("Calculation completed successfully.");
        }
        catch (Exception ex)
        {
            // Capture and display calculation errors
            Console.WriteLine("Calculation error captured:");
            Console.WriteLine(ex.Message);
        }

        try
        {
            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to save workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
