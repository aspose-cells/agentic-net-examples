// Title: Export an Excel workbook to JSON with camelCase property names using Aspose.Cells for .NET (C#)
// Common Searches: how to save an .xlsx workbook as camelCase JSON using Aspose.Cells in C# | Aspose.Cells JsonSaveOptions camelCase property naming example for .NET | convert Excel file to JSON with camelCase keys using C# Aspose.Cells library | C# code to export Excel to JSON with lowerCamelCase property names via Aspose.Cells | Aspose.Cells workbook.Save JSON camelCase output C# tutorial
// Tags: Aspose.Cells JsonSaveOptions camelCase naming | C# export Excel workbook to JSON | Excel to JSON conversion with camelCase keys | Aspose.Cells workbook.Save JSON output | JSON serialization of .xlsx using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, creates a JsonSaveOptions instance that uses camelCase property naming, and saves the workbook as a JSON file while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.json";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default settings)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Export the workbook to JSON using the configured options
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved as JSON to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
