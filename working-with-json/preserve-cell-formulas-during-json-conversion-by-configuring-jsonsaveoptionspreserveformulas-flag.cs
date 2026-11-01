// Title: Save an Excel workbook to JSON while preserving cell formulas using Aspose.Cells JsonSaveOptions in C#
// AI Prompts: Write C# code that loads an .xlsx file, sets JsonSaveOptions.PreserveFormulas = true, and saves the workbook as a .json file with Aspose.Cells. | Update a C# Aspose.Cells example to create missing output directories and enable formula preservation during JSON export.
// Common Searches: Aspose.Cells C# preserve formulas when exporting to JSON | JsonSaveOptions PreserveFormulas property usage example | How to keep Excel formulas in JSON output with Aspose.Cells | C# convert workbook to JSON with formulas retained | Saving Excel as JSON with formula retention using Aspose.Cells
// Tags: Aspose.Cells JsonSaveOptions PreserveFormulas | export Excel to JSON with formulas C# | configure JSON save options Aspose.Cells | C# workbook to JSON preserving formulas | Aspose.Cells formula retention during JSON conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an existing Excel file (or creates a new workbook), configures JsonSaveOptions.PreserveFormulas = true, ensures the output folder exists, and saves the workbook as a JSON file. Enabling PreserveFormulas retains any cell formulas in the generated JSON output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.json";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            }

            // Configure JSON save options (no ExportFormulas property in JsonSaveOptions)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as JSON with the configured options
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
