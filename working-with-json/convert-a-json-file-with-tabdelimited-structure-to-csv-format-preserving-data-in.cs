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
            // Path to the source JSON file (contains tab‑delimited data representation)
            string jsonFilePath = "input.json";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure JSON import options.
            // ArrayAsTable = true treats each array element as a separate row.
            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                // Preserve original field names as headers
                IgnoreTitle = false
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, cells, 0, 0, jsonOptions);

            // Save the workbook as CSV (comma‑separated values)
            string csvOutputPath = "output.csv";
            workbook.Save(csvOutputPath, SaveFormat.Csv);

            Console.WriteLine($"JSON data successfully converted to CSV at: {csvOutputPath}");
        }
    }
}