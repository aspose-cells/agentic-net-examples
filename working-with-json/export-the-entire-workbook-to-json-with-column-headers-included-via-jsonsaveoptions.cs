// Title: Export a full Excel workbook to JSON including column headers with Aspose.Cells JsonSaveOptions (C#)
// AI Prompts: Write a C# program that loads an .xlsx file, configures JsonSaveOptions to keep the first row as field names, and saves the workbook as a .json file using Aspose.Cells. | Show how to use Aspose.Cells Workbook.Save together with JsonSaveOptions to serialize all worksheets into a single JSON document while preserving header rows.
// Common Searches: C# Aspose.Cells export whole workbook to JSON with column headers | How to keep first row as keys when converting Excel to JSON using Aspose | Aspose.Cells JsonSaveOptions default includes headers example | Convert multi‑sheet Excel file to JSON preserving headers in .NET | Save Excel workbook as JSON file with Aspose.Cells library
// Tags: Aspose.Cells JsonSaveOptions JSON export | C# workbook.Save JSON with headers | Excel to JSON conversion Aspose.Cells | preserve column headers JSON export | full workbook serialization to JSON C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// // This C# example verifies that input.xlsx exists, loads it with Aspose.Cells, applies the default JsonSaveOptions (which retain column headers), and saves the entire workbook—including all worksheets—to output.json, with graceful error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default includes column headers)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Export the entire workbook to JSON
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved as JSON to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
