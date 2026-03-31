using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace OxpsJsonToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input OXPS JSON file path (replace with actual path)
            string jsonFilePath = "input.json";

            // Output CSV file path
            string csvFilePath = "output.csv";

            // Read the JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Set up JSON layout options (default options are sufficient for most cases)
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                // Process JSON arrays as tables so each array element becomes a row
                ArrayAsTable = true
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

            // Save the workbook as CSV
            workbook.Save(csvFilePath, SaveFormat.Csv);

            Console.WriteLine($"Conversion completed. CSV saved to: {csvFilePath}");
        }
    }
}