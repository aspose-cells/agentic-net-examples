// Title: Export an Excel workbook to JSON using Aspose.Cells for .NET with default JsonSaveOptions (C#)
// Common Searches: how to save an .xlsx file as .json using Aspose.Cells in C# | Aspose.Cells default JsonSaveOptions export all worksheets to JSON | C# convert Excel workbook to JSON with Aspose.Cells library | save workbook to JSON file using Aspose.Cells .NET API | export Excel data to JSON without custom settings Aspose.Cells
// Tags: Aspose.Cells JsonSaveOptions export | C# workbook.Save JSON | Aspose.Cells convert Excel to JSON | default JSON serialization Aspose.Cells | export Excel workbook as JSON .NET | Aspose.Cells save format JSON

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The program loads an Excel workbook from "input.xlsx", creates a JsonSaveOptions instance with default settings (which serializes all worksheets), and saves the workbook as a JSON file "activeWorksheet.json" using workbook.Save. It includes file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "activeWorksheet.json";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create JSON save options (default exports all worksheets)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
