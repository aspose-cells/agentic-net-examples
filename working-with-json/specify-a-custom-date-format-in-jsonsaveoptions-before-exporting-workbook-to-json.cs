// Title: How to specify a custom date format when saving an Excel workbook to JSON with Aspose.Cells for .NET
// Common Searches: Aspose.Cells JsonSaveOptions custom date format C# example | set date format for JSON export of Excel workbook using Aspose.Cells | C# export Excel to JSON with specific date pattern Aspose.Cells | JsonSaveOptions date formatting not applied Aspose.Cells .NET | how to change JSON date representation when saving workbook with Aspose.Cells | Aspose.Cells JSON save options date format limitation
// Tags: JsonSaveOptions date format customization | Aspose.Cells JSON workbook export | C# Excel to JSON conversion with date handling | JSON save options limitation for dates | Aspose.Cells JSON output date representation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel file into a Workbook, creates a JsonSaveOptions object (noting that custom date formatting is not supported in the current Aspose.Cells version), and saves the workbook as JSON while handling file‑not‑found and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (custom date format not supported in this version)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to JSON: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
