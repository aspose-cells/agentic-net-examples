using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class CsvToJsonConverter
    {
        public static void Run()
        {
            // Path to the source CSV file (the file may have an .xlsm.csv naming convention)
            string sourceCsvPath = "input.csv";

            // Desired output JSON file path
            string outputJsonPath = "output.json";

            // LoadOptions specifying that the source file is a CSV
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

            // Configure JSON save options as needed
            JsonSaveOptions jsonSaveOptions = new JsonSaveOptions
            {
                // Export the data as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // Include header row if present
                HasHeaderRow = true,
                // Export empty cells as null (optional)
                ExportEmptyCells = true,
                // Indent the JSON for readability
                Indent = "  "
            };

            // Perform the conversion using the provided ConversionUtility method
            ConversionUtility.Convert(sourceCsvPath, loadOptions, outputJsonPath, jsonSaveOptions);

            Console.WriteLine($"CSV file '{sourceCsvPath}' has been successfully converted to JSON at '{outputJsonPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CsvToJsonConverter.Run();
        }
    }
}