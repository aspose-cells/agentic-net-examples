// Title: Save an Excel worksheet as JSON with column headers as keys using Aspose.Cells for C#
// Common Searches: how to save Excel file to JSON using Aspose.Cells C# with column headers as property names | Aspose.Cells JsonSaveOptions IncludeColumnNames example in C# | convert .xlsx to .json preserving column names with Aspose.Cells | C# code to export workbook to JSON with headers as keys using Aspose.Cells
// Tags: Aspose.Cells JSON export with headers | C# export worksheet to JSON using JsonSaveOptions | include column names as JSON keys Aspose.Cells | convert Excel workbook to JSON format C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing .xlsx workbook, configures JsonSaveOptions to include column headers as JSON keys, and saves the data to a .json file, with error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default settings export column headers as keys)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook to a JSON file using the configured options
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
