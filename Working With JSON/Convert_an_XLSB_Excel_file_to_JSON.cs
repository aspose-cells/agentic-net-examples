using System;
using Aspose.Cells;

namespace AsposeCellsJsonConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSB file
            string sourcePath = "input.xlsb";

            // Path where the JSON output will be saved
            string jsonPath = "output.json";

            // Load the XLSB workbook
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single worksheet
                AlwaysExportAsJsonObject = true,

                // Include empty rows in the output (optional, adjust as needed)
                SkipEmptyRows = false,

                // Export nested structure (optional, adjust as needed)
                ExportNestedStructure = true
            };

            // Save the workbook as JSON using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {jsonPath}");
        }
    }
}