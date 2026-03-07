using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExcelToJsonConverter
    {
        /// <summary>
        /// Converts an Excel file to a JSON representation.
        /// </summary>
        /// <param name="excelPath">Full path to the source Excel file.</param>
        /// <param name="jsonPath">Full path where the JSON output will be saved.</param>
        public static void ConvertExcelToJson(string excelPath, string jsonPath)
        {
            // Load the workbook from the Excel file.
            Workbook workbook = new Workbook(excelPath);

            // Configure JSON save options.
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                ToExcelStruct = true,
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };

            // Save the workbook as JSON using the configured options.
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {jsonPath}");
        }

        // Example usage
        public static void Run()
        {
            string sourceExcel = "sample.xlsx";          // Path to the Excel file
            string destinationJson = "output.json";      // Desired JSON output path

            ConvertExcelToJson(sourceExcel, destinationJson);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ExcelToJsonConverter.Run();
        }
    }
}