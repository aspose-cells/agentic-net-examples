// Title: How to ensure UTF-8 encoding when exporting an Excel workbook to JSON with Aspose.Cells for .NET
// Common Searches: Aspose.Cells export Excel to JSON UTF-8 C# | C# set JSON encoding to UTF-8 using JsonSaveOptions Aspose | Does JsonSaveOptions allow specifying character set in Aspose.Cells | Save workbook as UTF-8 JSON with Aspose.Cells .NET example | Default character encoding for JSON output in Aspose.Cells
// Tags: Aspose.Cells JSON export UTF-8 | JsonSaveOptions character encoding | C# export Excel to JSON Aspose.Cells | Workbook.Save JSON options .NET | Aspose.Cells default JSON encoding

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an existing XLSX file into an Aspose.Cells Workbook, creates a JsonSaveOptions object (which uses UTF-8 encoding by default), and saves the workbook as a JSON file. It demonstrates that no explicit encoding setting is required because Aspose.Cells outputs JSON in UTF-8 automatically.
    class JsonExportExample
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure JSON save options (no Encoding property needed)
                JsonSaveOptions jsonOptions = new JsonSaveOptions();

                // Save the workbook as JSON
                workbook.Save(outputPath, jsonOptions);
                Console.WriteLine($"Workbook successfully saved as JSON to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
