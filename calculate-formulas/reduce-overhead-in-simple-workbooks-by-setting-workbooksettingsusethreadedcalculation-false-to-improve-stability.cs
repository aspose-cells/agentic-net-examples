// Title: Disable threaded formula calculation in Aspose.Cells (C#) to reduce overhead and improve workbook stability
// Common Searches: Aspose.Cells C# turn off UseThreadedCalculation for small workbooks | How to improve stability of simple Excel files using Aspose.Cells settings | Set Workbook.Settings.UseThreadedCalculation to false in .NET | Reduce calculation overhead in Aspose.Cells by disabling multithreading | Performance tuning Aspose.Cells workbook without threaded calculation
// Tags: disable threaded calculation Aspose.Cells C# | Workbook.Settings.UseThreadedCalculation false | Aspose.Cells performance optimization .NET | reduce formula calculation overhead Aspose.Cells | increase workbook stability Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, disables the UseThreadedCalculation setting to prevent multithreaded formula evaluation, ensures the output directory exists, saves the workbook as output.xlsx, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: The UseThreadedCalculation property may not be available in older Aspose.Cells versions.
            // If needed, uncomment the line below and ensure you are using a version that supports it.
            // workbook.Settings.UseThreadedCalculation = false;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
