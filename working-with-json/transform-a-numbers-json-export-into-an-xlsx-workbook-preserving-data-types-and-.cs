using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace NumbersJsonToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input JSON file exported from Apple Numbers
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "numbers_export.json");

            // If the JSON file does not exist, create a simple example JSON
            if (!File.Exists(jsonFilePath))
            {
                string sampleJson = @"{
                    ""Sheet1"": [
                        { ""Name"": ""Alice"", ""Score"": 95, ""Date"": ""2023-01-15"" },
                        { ""Name"": ""Bob"",   ""Score"": 88, ""Date"": ""2023-01-16"" }
                    ]
                }";
                File.WriteAllText(jsonFilePath, sampleJson);
            }

            // Output XLSX workbook path
            string outputXlsxPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "converted_output.xlsx");

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Get the first worksheet where data will be placed
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure layout options:
            // - Treat JSON arrays as tables
            // - Convert strings that represent numbers or dates to proper types
            // - Apply numeric and date formats to preserve appearance
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true,
                NumberFormat = "#,##0.00",
                DateFormat = "yyyy-MM-dd"
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the workbook as XLSX
            workbook.Save(outputXlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"JSON data successfully converted and saved to '{outputXlsxPath}'.");
        }
    }
}