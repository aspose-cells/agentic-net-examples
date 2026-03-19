using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace DifJsonToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input JSON file that contains DIF‑formatted data
            string jsonInputPath = "input.json";

            // Output XLSX file path
            string xlsxOutputPath = "output.xlsx";

            // Read the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonInputPath);

            // Create a new empty workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet where data will be imported
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure JSON import options (optional settings can be adjusted)
            JsonLayoutOptions importOptions = new JsonLayoutOptions
            {
                // Example: treat JSON arrays as tables; adjust as needed
                ArrayAsTable = true
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            // (feature rule: JsonUtility.ImportData)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, importOptions);

            // Save the workbook as XLSX (lifecycle: save)
            workbook.Save(xlsxOutputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. XLSX file saved to: {xlsxOutputPath}");
        }
    }
}