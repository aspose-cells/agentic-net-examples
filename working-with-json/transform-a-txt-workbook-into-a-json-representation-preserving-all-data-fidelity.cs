using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TxtToJsonConverter
    {
        public static void Run()
        {
            // Path to the source TXT file (can be CSV, tab‑delimited, etc.)
            string txtPath = "input.txt";

            // Load the TXT workbook. Aspose.Cells automatically detects the format.
            Workbook workbook = new Workbook(txtPath);

            // Configure JSON save options to preserve the full workbook structure.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a hierarchical JSON object.
                ExportNestedStructure = true,

                // Keep the Excel‑style structure (worksheets, tables, etc.) in JSON.
                ToExcelStruct = true,

                // Include empty cells as null to retain exact layout.
                ExportEmptyCells = true,

                // Export cell values as strings to avoid type loss.
                ExportAsString = true,

                // Ensure that even a single worksheet is wrapped in a JSON object.
                AlwaysExportAsJsonObject = true
            };

            // Save the workbook as a JSON file using the configured options.
            string jsonOutputPath = "output.json";
            workbook.Save(jsonOutputPath, jsonOptions);

            Console.WriteLine($"TXT workbook has been converted to JSON and saved to: {jsonOutputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TxtToJsonConverter.Run();
        }
    }
}