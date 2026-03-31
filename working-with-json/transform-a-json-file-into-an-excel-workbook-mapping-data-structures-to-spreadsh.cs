using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace JsonToExcelDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input JSON file path and output Excel file path
            string jsonFilePath = "input.json";
            string excelOutputPath = "output.xlsx";

            // Read the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook (empty spreadsheet)
            Workbook workbook = new Workbook();

            // Get the first worksheet where data will be imported
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure layout options for importing JSON.
            // Setting ArrayAsTable = true will map JSON arrays to tabular rows.
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                // Additional options can be set here if needed, e.g. IgnoreTitle, ConvertNumericOrDate, etc.
            };

            // Import the JSON string into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the populated workbook to the specified Excel file
            workbook.Save(excelOutputPath);

            Console.WriteLine($"JSON data from '{jsonFilePath}' has been successfully imported to '{excelOutputPath}'.");
        }
    }
}