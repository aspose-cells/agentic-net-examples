using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversion
{
    class Program
    {
        static void Main()
        {
            // Paths for the source JSON file (exported from an XLSX) and the target CSV file
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.json");
            string csvFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.csv");

            // Ensure the JSON file exists; if not, create a simple sample JSON
            if (!File.Exists(jsonFilePath))
            {
                string sampleJson = "[{\"Name\":\"John\",\"Age\":30},{\"Name\":\"Jane\",\"Age\":25}]";
                File.WriteAllText(jsonFilePath, sampleJson);
            }

            // Read the JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Prepare JSON layout options – treat JSON arrays as tables for proper cell mapping
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import the JSON data into the first worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

            // Save the populated workbook as CSV
            workbook.Save(csvFilePath, SaveFormat.Csv);

            Console.WriteLine($"Conversion completed. CSV saved to: {csvFilePath}");
        }
    }
}