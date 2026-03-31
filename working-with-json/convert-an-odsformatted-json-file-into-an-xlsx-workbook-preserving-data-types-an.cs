using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace OdsJsonToXlsx
{
    class Program
    {
        static void Main()
        {
            // Path to the JSON file that contains ODS‑style data
            string jsonFilePath = "input.json";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create an empty workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet where data will be imported
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure JSON import options to preserve data types and structure
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                // Treat JSON arrays as tables (each object becomes a row)
                ArrayAsTable = true,

                // Convert numeric strings and date strings to proper Excel types
                ConvertNumericOrDate = true,

                // Optional: specify expected date and number formats
                DateFormat = "yyyy-MM-dd",
                NumberFormat = "0.##"
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            // (lifecycle rule: load)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the workbook as XLSX (lifecycle rule: save)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Conversion completed: input.json -> output.xlsx");
        }
    }
}