using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace XpsJsonToCsvConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input JSON file (generated from XPS) and output CSV file paths
            string jsonFilePath = "input.json";
            string csvFilePath = "output.csv";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure JSON import options: treat arrays as tables for proper column alignment
            JsonLayoutOptions jsonOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, cells, 0, 0, jsonOptions);

            // Save the populated workbook as CSV, preserving column alignment
            workbook.Save(csvFilePath, SaveFormat.Csv);

            Console.WriteLine($"Conversion completed: '{jsonFilePath}' → '{csvFilePath}'");
        }
    }
}