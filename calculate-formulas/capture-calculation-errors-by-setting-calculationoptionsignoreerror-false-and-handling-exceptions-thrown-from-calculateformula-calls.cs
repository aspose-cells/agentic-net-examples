// Title: Capture Excel formula calculation errors in Aspose.Cells for .NET by disabling IgnoreError
// Common Searches: Aspose.Cells calculateformula throws exception when IgnoreError is false in C# | How to enable formula error exceptions in Aspose.Cells .NET | C# capture Excel calculation errors using Aspose.Cells CalculationOptions | Set CalculationOptions.IgnoreError to false to get detailed formula errors Aspose.Cells | Handling formula evaluation errors with Aspose.Cells workbook.CalculateFormula
// Tags: Aspose.Cells calculationoptions ignoreerror | Aspose.Cells calculateformula error handling | Excel formula exception capture C# | Aspose.Cells workbook calculation error handling | Aspose.Cells load or create workbook with formula

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file or creates a new workbook, configures CalculationOptions.IgnoreError to false, and runs workbook.CalculateFormula inside a try‑catch block to capture any formula evaluation exceptions, then saves the workbook while handling load and save errors.
class Program
{
    static void Main()
    {
        Workbook workbook = null;
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load workbook if the file exists; otherwise create a new workbook
        try
        {
            if (File.Exists(inputPath))
            {
                // Attempt to load the existing workbook
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                // Optionally add sample data/formulas here
                workbook.Worksheets[0].Cells["A1"].PutValue(10);
                workbook.Worksheets[0].Cells["A2"].PutValue(20);
                workbook.Worksheets[0].Cells["A3"].Formula = "SUM(A1:A2)";
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            // Fallback to a new workbook if loading fails
            workbook = new Workbook();
        }

        // Set calculation options to throw on errors
        CalculationOptions calcOptions = new CalculationOptions
        {
            IgnoreError = false
        };

        // Perform formula calculation with the specified options
        try
        {
            workbook.CalculateFormula(calcOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Calculation error: " + ex.Message);
        }

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
