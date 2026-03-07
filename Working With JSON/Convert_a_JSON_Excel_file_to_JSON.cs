using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an Excel workbook
            string sourceJsonPath = "input.json";

            // Load the JSON file into a Workbook object.
            // JsonLoadOptions can be customized if needed; using defaults here.
            JsonLoadOptions loadOptions = new JsonLoadOptions();
            Workbook workbook = new Workbook(sourceJsonPath, loadOptions);

            // Configure options for saving the workbook as JSON.
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if it contains a single sheet.
                AlwaysExportAsJsonObject = true,
                // Convert the Excel structure to a JSON struct (hierarchical representation).
                ToExcelStruct = true,
                // Optional: include header rows, export empty cells, etc., as required.
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportAsString = false,
                SkipEmptyRows = true
            };

            // Save the workbook back to JSON format.
            string outputJsonPath = "output.json";
            workbook.Save(outputJsonPath, saveOptions);

            // Indicate completion.
            Console.WriteLine($"JSON conversion completed. Output saved to: {outputJsonPath}");
        }
    }
}