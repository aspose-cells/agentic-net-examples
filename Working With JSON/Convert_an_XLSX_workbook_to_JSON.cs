using System;
using Aspose.Cells;
using Aspose.Cells.Json; // Namespace for JsonSaveOptions (if needed)

namespace AsposeCellsExamples
{
    public class XlsxToJsonConverter
    {
        /// <summary>
        /// Converts an existing XLSX workbook to a JSON file.
        /// </summary>
        /// <param name="inputPath">Full path of the source XLSX file.</param>
        /// <param name="outputPath">Full path where the JSON file will be saved.</param>
        public static void ConvertXlsxToJson(string inputPath, string outputPath)
        {
            // Load the workbook from the specified XLSX file.
            // Uses the Workbook(string) constructor as per the lifecycle rule.
            Workbook workbook = new Workbook(inputPath);

            // Configure JSON save options.
            // Example: always export as a JSON object and include empty cells.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportAsString = true,
                Indent = "  "
            };

            // Save the workbook as JSON using the Save(string, SaveOptions) method.
            // This follows the provided save rule.
            workbook.Save(outputPath, jsonOptions);

            // Optional: inform the user.
            Console.WriteLine($"Workbook '{inputPath}' has been successfully converted to JSON at '{outputPath}'.");
        }

        // Example usage
        public static void Main()
        {
            string sourceXlsx = "sample.xlsx";
            string destinationJson = "sample.json";

            ConvertXlsxToJson(sourceXlsx, destinationJson);
        }
    }
}