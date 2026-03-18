using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTransformation
{
    public class XmlJsonToCsvConverter
    {
        /// <summary>
        /// Converts an XML or JSON file to CSV while preserving data types.
        /// </summary>
        /// <param name="sourcePath">Full path of the source XML or JSON file.</param>
        /// <param name="csvPath">Full path where the resulting CSV file will be saved.</param>
        public static void ConvertToCsv(string sourcePath, string csvPath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
                throw new ArgumentException("Source path must be provided.", nameof(sourcePath));

            if (string.IsNullOrWhiteSpace(csvPath))
                throw new ArgumentException("CSV path must be provided.", nameof(csvPath));

            // Determine load format based on file extension
            string ext = Path.GetExtension(sourcePath).ToLowerInvariant();
            LoadOptions loadOptions;

            switch (ext)
            {
                case ".xml":
                    // Use XmlLoadOptions to correctly interpret numeric/date values
                    var xmlOptions = new XmlLoadOptions(LoadFormat.Xml);
                    xmlOptions.ConvertNumericOrDate = true;
                    loadOptions = xmlOptions;
                    break;

                case ".json":
                    // Json is treated as a generic load format; Aspose.Cells can load it directly
                    loadOptions = new LoadOptions(LoadFormat.Json);
                    break;

                default:
                    throw new NotSupportedException($"File extension '{ext}' is not supported. Use .xml or .json.");
            }

            // Load the source file into a workbook using the determined options
            var workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook as CSV. SaveFormat.Csv ensures proper delimiter handling.
            workbook.Save(csvPath, SaveFormat.Csv);
        }

        // Example usage
        public static void Run()
        {
            try
            {
                // Example XML to CSV
                string xmlSource = "data.xml";
                string xmlCsv = "data_from_xml.csv";
                ConvertToCsv(xmlSource, xmlCsv);
                Console.WriteLine($"XML converted to CSV: {xmlCsv}");

                // Example JSON to CSV
                string jsonSource = "data.json";
                string jsonCsv = "data_from_json.csv";
                ConvertToCsv(jsonSource, jsonCsv);
                Console.WriteLine($"JSON converted to CSV: {jsonCsv}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XmlJsonToCsvConverter.Run();
        }
    }
}