using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (exported from Excel)
            string jsonFilePath = "source.json";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet where the data will be imported
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure import options.
            // Setting ArrayAsTable = true treats a JSON array as a table,
            // which is the typical format for Excel‑style JSON.
            JsonLayoutOptions importOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, importOptions);

            // Save the workbook as an Excel file
            string outputExcelPath = "output.xlsx";
            workbook.Save(outputExcelPath);

            Console.WriteLine($"JSON data has been converted and saved to '{outputExcelPath}'.");
        }
    }
}