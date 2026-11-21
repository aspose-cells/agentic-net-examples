// Title: How to save an Excel workbook to JSON while omitting hidden rows and columns with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export Excel to JSON without hidden rows | How to ignore hidden columns when saving workbook as JSON using JsonSaveOptions | Saving Excel as JSON while skipping hidden rows and columns in C# | Configure JsonSaveOptions to omit hidden data in JSON output | Export Excel data to JSON and exclude hidden cells with Aspose.Cells
// Tags: Aspose.Cells JsonSaveOptions hide rows exclusion | Excel to JSON export omit hidden columns | C# JsonSaveOptions filter out hidden data | Aspose.Cells export JSON without hidden elements | JSON save options for removing hidden rows

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

// The example loads an Excel file into an Aspose.Cells Workbook, creates a JsonSaveOptions instance (which by default excludes hidden rows and columns), and saves the workbook as a JSON file using those options in C#.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options (default behavior excludes hidden rows/columns)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON using the configured options
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to JSON: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
