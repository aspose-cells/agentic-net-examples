using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonFlattenDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that contains nested objects
            string inputJsonPath = "nested_input.json";

            // Path for the flattened JSON output
            string outputJsonPath = "flattened_output.json";

            // Load the JSON file into a workbook.
            // JsonLoadOptions can be left with default settings; we only need to read the data.
            JsonLoadOptions loadOptions = new JsonLoadOptions();
            Workbook workbook = new Workbook(inputJsonPath, loadOptions);

            // Configure JSON save options to produce a flattened structure.
            // Setting ExportNestedStructure to false disables the parent‑child hierarchy.
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                ExportNestedStructure = false,   // flatten the JSON
                AlwaysExportAsJsonObject = true, // ensure the result is a JSON object even with one sheet
                SkipEmptyRows = true,            // optional: omit empty rows for cleaner output
                Indent = "    "                  // optional: pretty‑print with 4‑space indentation
            };

            // Save the workbook as a flattened JSON file.
            workbook.Save(outputJsonPath, saveOptions);

            Console.WriteLine($"Flattened JSON saved to: {outputJsonPath}");
        }
    }
}