// Title: Create a C# class derived from AbstractCalculationMonitor to log circular reference cell addresses during formula calculation with Aspose.Cells
// Common Searches: how to capture circular reference cell addresses with Aspose.Cells calculation monitor in C# | C# Aspose.Cells custom AbstractCalculationMonitor example for logging circular references | override OnCircular method in Aspose.Cells to get circular reference details | detect and log circular references while calculating formulas using Aspose.Cells .NET | Aspose.Cells calculate formula monitor circular reference logging tutorial
// Tags: Aspose.Cells custom calculation monitor | log circular reference cells C# | AbstractCalculationMonitor OnCircular override | formula calculation circular reference detection .NET | Excel workbook formula evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, runs CalculateFormula, and saves the file. To obtain detailed information about circular references, developers should create a class that inherits from AbstractCalculationMonitor, override the OnCircular method to log the addresses of the cells involved, assign this monitor to the workbook's CalculationOptions, and then invoke CalculateFormula. This approach enables custom handling of circular reference errors during formula evaluation in Aspose.Cells for .NET.
public class Program
{
    public static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Perform calculation. Any circular references will raise an exception that we catch below.
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Calculation error: {calcEx.Message}");
                // Continue to save the workbook if needed.
            }

            // Save the workbook after calculation (optional).
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
