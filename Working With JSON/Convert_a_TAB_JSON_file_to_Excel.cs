using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the source JSON file (tabular JSON) and the destination Excel file
            string jsonFilePath = "input.json";
            string excelFilePath = "output.xlsx";

            // Read the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook (in-memory) and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure JSON layout options to treat JSON arrays as tables
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true   // Each array element becomes a row in the worksheet
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the populated workbook as an Excel file (XLSX format)
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Excel file saved to: {excelFilePath}");
        }
    }
}