// Title: How to export an Excel workbook to indented (pretty‑printed) JSON using Aspose.Cells for .NET (C#)
// Common Searches: C# code to save an Excel file as pretty printed JSON with Aspose.Cells | Aspose.Cells JsonSaveOptions indentation example for workbook export | Export Excel workbook to formatted JSON file using .NET SDK | How to generate indented JSON from an .xlsx workbook in C# | Saving workbook to JSON with readable formatting using Aspose.Cells
// Tags: JsonSaveOptions pretty print Aspose.Cells | Excel workbook JSON export .NET | formatted JSON output from workbook | save workbook as indented JSON | Aspose.Cells Excel to JSON conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, uses JsonSaveOptions (which default to readable formatting) to export the workbook, and saves the result as output.json with pretty‑printed indentation, handling errors and reporting success.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default formatting is already readable)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Export the workbook to a JSON file using the configured options
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
