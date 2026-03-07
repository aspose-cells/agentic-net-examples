using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPrnToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PRN/CSV file
            string prnFilePath = "input.prn";

            // Ensure the input file exists; create a sample if it does not
            if (!File.Exists(prnFilePath))
            {
                File.WriteAllText(prnFilePath, "Name,Age,City\nJohn,30,New York\nJane,25,London");
            }

            // Path where the resulting JSON will be saved
            string jsonOutputPath = "output.json";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the PRN (CSV) data into the worksheet starting at cell A1 (row 0, column 0)
            cells.ImportCSV(prnFilePath, ",", true, 0, 0);

            // Determine the used range dimensions
            int totalRows = cells.MaxRow + 1;      // MaxRow is zero‑based
            int totalColumns = cells.MaxColumn + 1;

            // Create a range that covers all imported data
            AsposeRange dataRange = cells.CreateRange(0, 0, totalRows, totalColumns);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportNestedStructure = true,
                ExportEmptyCells = true
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            // Write the JSON string to a file
            File.WriteAllText(jsonOutputPath, jsonResult);

            // Output the JSON to console for verification
            Console.WriteLine("JSON conversion completed. Result:");
            Console.WriteLine(jsonResult);
        }
    }
}