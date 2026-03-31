using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the HTML file that contains JSON data
            string htmlFilePath = "input.html";

            // Read the entire HTML content
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Simple extraction of the JSON part (assumes the first '{' starts the JSON and the last '}' ends it)
            int jsonStart = htmlContent.IndexOf('{');
            int jsonEnd = htmlContent.LastIndexOf('}');
            if (jsonStart == -1 || jsonEnd == -1 || jsonEnd <= jsonStart)
            {
                Console.WriteLine("Unable to locate JSON data within the HTML file.");
                return;
            }

            string jsonData = htmlContent.Substring(jsonStart, jsonEnd - jsonStart + 1);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set JSON layout options to treat arrays as tables and convert numbers/dates
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true
            };

            // Import the extracted JSON into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonData, cells, 0, 0, layoutOptions);

            // Save the workbook as CSV, preserving the field mappings
            string csvOutputPath = "output.csv";
            workbook.Save(csvOutputPath, SaveFormat.Csv);

            Console.WriteLine($"JSON data extracted from HTML and saved as CSV to '{csvOutputPath}'.");
        }
    }
}