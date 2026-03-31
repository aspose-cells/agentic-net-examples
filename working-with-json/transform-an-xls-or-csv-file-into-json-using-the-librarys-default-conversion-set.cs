using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonConversion
{
    public class Converter
    {
        /// <summary>
        /// Converts an Excel (XLS/XLSX) or CSV file to JSON using Aspose.Cells default settings.
        /// </summary>
        /// <param name="sourcePath">Full path of the source XLS or CSV file.</param>
        /// <param name="jsonPath">Full path where the resulting JSON file will be saved.</param>
        public static void ConvertToJson(string sourcePath, string jsonPath)
        {
            // Load the workbook. Aspose.Cells automatically detects the format (XLS, XLSX, CSV, etc.).
            Workbook workbook = new Workbook(sourcePath);

            // Create default JSON save options. No custom settings are applied, so default conversion is used.
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file.
            workbook.Save(jsonPath, jsonOptions);
        }

        // Example usage
        public static void Main()
        {
            // Example source files (replace with actual paths as needed)
            string xlsSource = "sample.xls";
            string csvSource = "sample.csv";

            // Destination JSON files
            string xlsJson = "sample_from_xls.json";
            string csvJson = "sample_from_csv.json";

            // Convert XLS to JSON
            ConvertToJson(xlsSource, xlsJson);
            Console.WriteLine($"Converted {xlsSource} to {xlsJson}");

            // Convert CSV to JSON
            ConvertToJson(csvSource, csvJson);
            Console.WriteLine($"Converted {csvSource} to {csvJson}");
        }
    }
}