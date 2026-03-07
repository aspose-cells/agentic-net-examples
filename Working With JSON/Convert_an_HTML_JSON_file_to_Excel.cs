using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class HtmlJsonToExcelConverter
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source JSON file (containing HTML‑derived JSON data)
            string jsonFilePath = "input.json";

            // Read the JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook (empty Excel file)
            Workbook workbook = new Workbook();

            // Get the cells collection of the first worksheet
            Cells cells = workbook.Worksheets[0].Cells;

            // Configure JSON import options
            JsonLayoutOptions importOptions = new JsonLayoutOptions
            {
                // Treat a JSON array as a table so that each element becomes a row
                ArrayAsTable = true
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, cells, 0, 0, importOptions);

            // Save the populated workbook to an Excel file
            string outputExcelPath = "output.xlsx";
            workbook.Save(outputExcelPath);

            Console.WriteLine($"Conversion completed. Excel file saved to '{outputExcelPath}'.");
        }
    }
}