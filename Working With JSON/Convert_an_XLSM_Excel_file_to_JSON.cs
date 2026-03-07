using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XlsmToJsonConverter
    {
        /// <summary>
        /// Converts an XLSM workbook to a JSON file.
        /// </summary>
        /// <param name="sourcePath">Full path of the source XLSM file.</param>
        /// <param name="jsonPath">Full path where the JSON output will be saved.</param>
        public static void ConvertXlsmToJson(string sourcePath, string jsonPath)
        {
            // Load the XLSM workbook from the specified file.
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options.
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false,
                ExportAsString = false
            };

            // Save the workbook as JSON using the configured options.
            workbook.Save(jsonPath, saveOptions);
        }

        // Example usage
        public static void Run()
        {
            string sourceFile = @"C:\Data\sample.xlsm";
            string outputFile = @"C:\Data\sample.json";

            try
            {
                ConvertXlsmToJson(sourceFile, outputFile);
                Console.WriteLine($"Conversion completed. JSON saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            XlsmToJsonConverter.Run();
        }
    }
}