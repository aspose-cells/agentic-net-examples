// Title: Export the active worksheet of an Excel workbook to JSON with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, identifies the active worksheet, copies it into a new workbook, and saves the result as a JSON file using Aspose.Cells default SaveFormat. | Show how to add file‑existence validation and try‑catch error handling when exporting the active sheet of an Excel workbook to JSON with Aspose.Cells. | Demonstrate copying only the active worksheet to a temporary workbook and invoking Workbook.Save to produce a JSON output without affecting other sheets.
// Common Searches: C# Aspose.Cells export only the active sheet to JSON file | How to save an Excel worksheet as JSON using Aspose.Cells .NET | Aspose.Cells copy active worksheet to new workbook before JSON conversion | Check if input Excel file exists before exporting to JSON with Aspose.Cells | Default SaveFormat Json example in Aspose.Cells C#
// Tags: export active worksheet to JSON Aspose.Cells | Workbook.Save with SaveFormat.Json C# | copy active sheet to temporary workbook Aspose.Cells | file existence validation Aspose.Cells | exception handling Aspose.Cells export

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input Excel file exists, loads it with Aspose.Cells, retrieves the active worksheet, copies that sheet into a new temporary workbook, and saves the temporary workbook as a JSON file using the default SaveFormat, while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the index of the active worksheet
            int activeWorksheetIndex = workbook.Worksheets.ActiveSheetIndex;
            string activeSheetName = workbook.Worksheets[activeWorksheetIndex].Name;

            // Create a temporary workbook containing only the active worksheet
            Workbook tempWorkbook = new Workbook();
            tempWorkbook.Worksheets.Clear(); // remove default sheet
            // Add a copy of the active worksheet by name (compatible overload)
            tempWorkbook.Worksheets.AddCopy(activeSheetName);

            // Save the active worksheet to a JSON file
            tempWorkbook.Save(outputPath, SaveFormat.Json);
            Console.WriteLine($"Active worksheet exported to JSON successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
