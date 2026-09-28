// Title: Export a complete Excel workbook to JSON with column headers using Aspose.Cells JsonSaveOptions in C#
// AI Prompts: Generate C# code that loads an .xlsx workbook, configures JsonSaveOptions to retain column headers, and saves the entire workbook as a JSON file with Aspose.Cells. | Show how to use Aspose.Cells JsonSaveOptions in .NET to export all worksheets of an Excel file to a single JSON document while preserving header rows.
// Common Searches: how to export all sheets from an Excel file to JSON with headers using Aspose.Cells C# | Aspose.Cells JsonSaveOptions include column names when saving workbook to JSON | C# convert multi-sheet Excel workbook to JSON preserving column headers Aspose
// Tags: Aspose.Cells JsonSaveOptions export workbook to JSON | C# export Excel to JSON with column headers | save entire workbook as JSON using Aspose.Cells | include header rows in JSON output Aspose.Cells | convert multi-sheet Excel to JSON .NET

using System;
using System.IO;
using Aspose.Cells;

// The program verifies the presence of the source Excel file, loads it with Aspose.Cells, creates a JsonSaveOptions instance (which includes column headers by default), and saves the entire workbook—including all worksheets—to a JSON file, handling any runtime exceptions.
class ExportWorkbookToJson
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.json";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default behavior includes column headers)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Export the entire workbook to a JSON file
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
