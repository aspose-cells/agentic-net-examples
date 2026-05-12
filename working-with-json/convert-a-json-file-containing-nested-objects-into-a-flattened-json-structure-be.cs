using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonFlattenDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source JSON file containing nested objects
            string inputJsonPath = "nested_input.json";

            // Path for the flattened JSON output
            string outputJsonPath = "flattened_output.json";

            // Load the JSON file into a workbook.
            // JsonLoadOptions can be used to keep the original schema if needed.
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                KeptSchema = true          // Preserve schema during load (optional)
            };

            // The constructor loads the JSON data into the workbook.
            Workbook workbook = new Workbook(inputJsonPath, loadOptions);

            // Configure JSON save options for a flattened structure.
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // When ExportNestedStructure is false, the JSON is saved as a flat table.
                ExportNestedStructure = false,

                // Export as a JSON object even if there is only one worksheet.
                AlwaysExportAsJsonObject = true,

                // Include header row if the first row contains column names.
                HasHeaderRow = true,

                // Skip empty rows to keep the output concise.
                SkipEmptyRows = true,

                // Export empty cells as null (optional, can be set to false).
                ExportEmptyCells = true
            };

            // Save the workbook as a flattened JSON file.
            workbook.Save(outputJsonPath, saveOptions);

            // Optional: display the resulting JSON content.
            Console.WriteLine("Flattened JSON saved to: " + outputJsonPath);
            Console.WriteLine(File.ReadAllText(outputJsonPath));
        }
    }
}