using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsToJsonConverter
    {
        /// <summary>
        /// Converts an XLS workbook to a JSON file.
        /// </summary>
        /// <param name="sourceXlsPath">Full path of the source XLS file.</param>
        /// <param name="outputJsonPath">Full path where the JSON file will be saved.</param>
        public static void Convert(string sourceXlsPath, string outputJsonPath)
        {
            // Load the XLS workbook from the specified file.
            // Using the Workbook(string) constructor as defined in the provided rules.
            Workbook workbook = new Workbook(sourceXlsPath);

            // Configure JSON save options.
            // The default options are sufficient for a basic conversion,
            // but you can customize them (e.g., ExportNestedStructure, HasHeaderRow, etc.).
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Example customizations (optional):
                // ExportNestedStructure = true,
                // HasHeaderRow = true,
                // ExportEmptyCells = true,
                // AlwaysExportAsJsonObject = true
            };

            // Save the workbook as a JSON file using the Save(string, SaveOptions) overload.
            // This follows the lifecycle rule for saving.
            workbook.Save(outputJsonPath, jsonOptions);
        }

        // Example usage
        public static void Main()
        {
            string sourcePath = "input.xls";      // Path to the source XLS workbook
            string jsonPath   = "output.json";    // Desired output JSON file path

            Convert(sourcePath, jsonPath);
            Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
        }
    }
}