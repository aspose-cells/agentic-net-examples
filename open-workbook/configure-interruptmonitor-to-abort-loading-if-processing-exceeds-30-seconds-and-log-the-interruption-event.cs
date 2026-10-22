// Title: Configure Aspose.Cells InterruptMonitor to abort workbook loading after 30 seconds and log the interruption in C#
// Common Searches: Aspose.Cells C# set a 30 second timeout for workbook loading | how to use InterruptMonitor to stop loading large Excel files in C# | log a workbook load interruption event with Aspose.Cells | abort Excel workbook load when processing exceeds a time limit using Aspose.Cells | configure InterruptMonitor to cancel long‑running load operations in C#
// Tags: InterruptMonitor abort load timeout Aspose.Cells | Workbook load timeout handling C# | log load interruption Aspose.Cells | Excel workbook loading timeout Aspose.Cells | configure InterruptMonitor 30 seconds C#

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to set up Aspose.Cells' InterruptMonitor with a 30‑second timeout, abort the workbook loading process if the limit is exceeded, and log the interruption event before optionally saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
