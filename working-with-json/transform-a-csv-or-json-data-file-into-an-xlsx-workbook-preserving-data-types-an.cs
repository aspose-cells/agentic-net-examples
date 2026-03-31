using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    /// <summary>
    /// Utility class to convert CSV or JSON files to XLSX while preserving data types.
    /// </summary>
    public static class CsvJsonToXlsxConverter
    {
        /// <summary>
        /// Converts the specified source file (CSV or JSON) to an XLSX workbook.
        /// </summary>
        /// <param name="sourcePath">Full path of the input CSV or JSON file.</param>
        /// <param name="destPath">Full path where the resulting XLSX file will be saved.</param>
        public static void ConvertToXlsx(string sourcePath, string destPath)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentException("Source path must be provided.", nameof(sourcePath));

            if (string.IsNullOrEmpty(destPath))
                throw new ArgumentException("Destination path must be provided.", nameof(destPath));

            // Determine the load format based on file extension
            LoadFormat loadFormat;
            string extension = Path.GetExtension(sourcePath).ToLowerInvariant();

            switch (extension)
            {
                case ".csv":
                    loadFormat = LoadFormat.Csv;
                    break;
                case ".json":
                    loadFormat = LoadFormat.Json;
                    break;
                default:
                    throw new NotSupportedException($"File extension '{extension}' is not supported. Only .csv and .json are allowed.");
            }

            // Create load options for the detected format
            LoadOptions loadOptions = new LoadOptions(loadFormat);

            // Load the source file into a Workbook instance
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook as XLSX (preserves numeric, date, and string types)
            workbook.Save(destPath, SaveFormat.Xlsx);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Example CSV conversion
            string csvSource = "sample.csv";
            string csvDest = "sample_converted.xlsx";

            // Example JSON conversion
            string jsonSource = "sample.json";
            string jsonDest = "sample_converted_from_json.xlsx";

            // Ensure sample files exist (in real scenarios they would already be present)
            File.WriteAllText(csvSource, "Name,Age,JoinDate\nJohn,30,2021-05-01\nAlice,25,2022-01-15");
            File.WriteAllText(jsonSource, "[{\"Name\":\"Bob\",\"Age\":40,\"JoinDate\":\"2020-12-31\"}]");

            // Perform conversions
            CsvJsonToXlsxConverter.ConvertToXlsx(csvSource, csvDest);
            CsvJsonToXlsxConverter.ConvertToXlsx(jsonSource, jsonDest);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}