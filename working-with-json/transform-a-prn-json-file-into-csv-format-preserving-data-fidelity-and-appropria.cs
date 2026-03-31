using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main()
        {
            // Path to the source PRN JSON file
            string jsonFilePath = "input.prn.json";

            if (!File.Exists(jsonFilePath))
            {
                Console.WriteLine($"File not found: {jsonFilePath}");
                return;
            }

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure JSON import options:
            // - ArrayAsTable = true to treat JSON arrays as tabular data
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the worksheet as CSV, preserving the imported data layout
            string csvOutputPath = "output.csv";
            workbook.Save(csvOutputPath, SaveFormat.Csv);

            Console.WriteLine($"JSON data from '{jsonFilePath}' has been converted to CSV at '{csvOutputPath}'.");
        }
    }
}