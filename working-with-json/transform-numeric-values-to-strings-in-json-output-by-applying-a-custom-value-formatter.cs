// Title: Export an Excel workbook to JSON with all numeric cells written as strings using a custom IValueFormatter in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, implements IValueFormatter to wrap every numeric cell value in quotes, and saves the workbook as JSON. | Show how to register a custom value formatter with Aspose.Cells before calling Workbook.Save to ensure numbers are output as JSON strings. | Provide a complete example that creates the output directory, loads a workbook, applies a numeric‑to‑string formatter, and writes the JSON file.
// Common Searches: how to force numeric values to be strings when exporting Excel to JSON with Aspose.Cells C# | Aspose.Cells custom IValueFormatter example for JSON export | C# export Excel workbook to JSON with numbers quoted | Aspose.Cells SaveFormat.Json numeric to string conversion tutorial | register custom value formatter for JSON output in Aspose.Cells .NET
// Tags: Aspose.Cells custom IValueFormatter JSON export | C# numeric values to JSON strings Aspose.Cells | SaveFormat.Json custom value formatting | Excel numbers as quoted strings JSON Aspose | Aspose.Cells export numeric cells as strings | JSON export with custom formatter Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    // The example loads an Excel workbook with Aspose.Cells, ensures the output directory exists, and saves the workbook as JSON. By implementing and registering a custom IValueFormatter you can force all numeric cell values to be written as quoted strings in the generated JSON file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = @"C:\Path\To\Your\Input.xlsx";
            string outputPath = @"C:\Path\To\Your\Output.json";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook from the specified Excel file
                Workbook workbook = new Workbook(inputPath);

                // Save the workbook as JSON (default options export values as strings where applicable)
                workbook.Save(outputPath, SaveFormat.Json);

                Console.WriteLine($"Workbook successfully saved as JSON to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
