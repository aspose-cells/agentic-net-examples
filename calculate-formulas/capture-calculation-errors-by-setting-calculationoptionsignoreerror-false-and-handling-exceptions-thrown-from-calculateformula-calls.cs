// Title: How to detect and handle Excel formula calculation errors in Aspose.Cells for .NET by disabling IgnoreError
// Common Searches: Aspose.Cells C# calculate formula without ignoring errors | Capture calculation exceptions when using Workbook.CalculateFormula in .NET | Set CalculationOptions.IgnoreError false to get formula errors in Aspose.Cells | Handle Excel formula errors programmatically with Aspose.Cells C# example | Prevent silent formula failures in Aspose.Cells by disabling error ignoring
// Tags: Aspose.Cells calculationoptions ignoreerror false | C# workbook calculateformula exception handling | Excel formula error detection Aspose.Cells | disable silent formula errors Aspose.Cells | Aspose.Cells workbook save after calculation errors

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, configures CalculationOptions with IgnoreError set to false, runs Workbook.CalculateFormula inside a try‑catch block to capture any formula calculation exceptions, logs the error message, and then saves the workbook while handling possible I/O errors.
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
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Configure calculation options (do not ignore errors)
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
            // Capture and handle calculation errors
            Console.WriteLine($"Calculation error: {ex.Message}");
        }

        try
        {
            // Save the workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
