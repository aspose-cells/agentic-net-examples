using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvFromJson
{
    class Program
    {
        static void Main()
        {
            // Path to the JSON file that contains CSV‑formatted data
            string jsonFilePath = "input.json";

            // Read the entire JSON content
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure JSON layout options:
            // - Treat JSON arrays as tables (rows)
            // - Convert numeric and date strings to proper Excel types
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                ConvertNumericOrDate = true
            };

            // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the workbook temporarily as XLSX (required for conversion utility)
            string tempXlsxPath = "temp.xlsx";
            workbook.Save(tempXlsxPath, SaveFormat.Xlsx);

            // Define the final CSV output path
            string csvOutputPath = "output.csv";

            // Convert the temporary XLSX file to CSV using the ConversionUtility
            ConversionUtility.Convert(tempXlsxPath, csvOutputPath);

            // Optional: clean up the temporary XLSX file
            if (File.Exists(tempXlsxPath))
            {
                File.Delete(tempXlsxPath);
            }

            Console.WriteLine($"CSV file has been created at: {Path.GetFullPath(csvOutputPath)}");
        }
    }
}