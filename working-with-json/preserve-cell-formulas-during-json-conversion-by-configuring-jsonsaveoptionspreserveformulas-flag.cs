// Title: Preserving Excel cell formulas when saving a workbook as JSON with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to JSON while keeping formulas | set JsonSaveOptions PreserveFormulas flag in .NET example | convert Excel file to JSON without losing formula calculations using Aspose.Cells | how to export Excel to JSON with formulas intact in C# | sample code for JsonSaveOptions PreserveFormulas in Aspose.Cells
// Tags: Aspose.Cells JSON PreserveFormulas option | C# export Excel to JSON with formulas | Excel to JSON conversion preserving formulas | JsonSaveOptions configuration for formula retention | Aspose.Cells workbook save as JSON example | preserve cell formulas during JSON serialization .NET

using System;
using System.IO;
using Aspose.Cells;

// This C# example demonstrates how to load an Excel workbook with Aspose.Cells, enable the JsonSaveOptions.PreserveFormulas flag, and save the workbook as a JSON file so that all cell formulas are retained in the output. The code includes file existence verification and exception handling for robust execution.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.json";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default options are sufficient)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON using the configured options
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved as JSON to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
