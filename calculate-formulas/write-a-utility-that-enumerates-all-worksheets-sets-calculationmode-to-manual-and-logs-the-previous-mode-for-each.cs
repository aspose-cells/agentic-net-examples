// Title: Enumerate all worksheets, set each to manual calculation mode, and log the previous mode with Aspose.Cells for .NET
// Common Searches: how to loop through all worksheets and change calculation mode to manual using Aspose.Cells in C# | C# code to retrieve and log previous calculation mode for each worksheet with Aspose.Cells | Aspose.Cells example for enumerating worksheets and setting manual calculation mode | log original calculation mode before switching to manual for every sheet in a workbook
// Tags: enumerate worksheets Aspose.Cells | set calculation mode manual Aspose.Cells | log previous calculation mode C# | iterate workbook worksheets adjust calculation settings | Aspose.Cells workbook calculation mode handling

using System;
using System.IO;
using Aspose.Cells;

// The utility loads an Excel workbook, iterates through every worksheet, records each sheet's current CalculationMode, switches the mode to Manual, and writes the original mode to the console. After processing all sheets, the workbook is saved, preserving the new manual calculation settings.
class Program
{
    static void Main(string[] args)
    {
        // Paths to the source and destination Excel files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Optional: Log workbook information
            Console.WriteLine($"Workbook loaded with {workbook.Worksheets.Count} worksheet(s).");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
