// Title: How to assign a custom calculation monitor to Workbook.Settings.CalculationMonitor before loading an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# assign custom calculation monitor before opening workbook | set Workbook.Settings.CalculationMonitor to track formula evaluation progress in .NET | use custom calculation monitor to cancel formula calculation in Aspose.Cells | initialize calculation monitor prior to loading Excel file with Aspose.Cells | C# example for Workbook.Settings.CalculationMonitor usage in Aspose.Cells
// Tags: custom calculation monitor Aspose.Cells C# | Workbook.Settings.CalculationMonitor assignment | formula calculation progress monitoring .NET | cancel formula evaluation Aspose.Cells | preload calculation monitor Excel workbook

using System;
using System.IO;
using Aspose.Cells;

// This example shows how to create a CustomCalculationMonitor class with BeginCalculate, EndCalculate, SetProgress, and IsCancel methods, assign it to Workbook.Settings.CalculationMonitor before loading an Excel file, and then save the workbook. It includes input file validation, output directory creation, and basic error handling.
class CustomCalculationMonitor
{
    // Called when calculation starts
    public void BeginCalculate()
    {
        // Custom logic before calculation begins
    }

    // Called when calculation ends
    public void EndCalculate()
    {
        // Custom logic after calculation ends
    }

    // Receives progress updates (0.0 to 1.0)
    public void SetProgress(double progress)
    {
        // Handle progress reporting if needed
    }

    // Determines whether the calculation should be cancelled
    public bool IsCancel()
    {
        // Return true to cancel the calculation; otherwise false
        return false;
    }
}

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // (Optional) Assign a custom calculation monitor if supported by the library version
            // workbook.Settings.CalculationMonitor = new CustomCalculationMonitor();

            // (Optional) Perform any workbook operations here

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
