using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace NumbersJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Determine the path to the Numbers JSON file relative to the executable
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "numbers.json");

            // If the JSON file does not exist, create a simple sample file
            if (!File.Exists(jsonFilePath))
            {
                string sampleJson = @"{
    ""data"": [
        { ""Name"": ""Alice"", ""Score"": 95 },
        { ""Name"": ""Bob"",   ""Score"": 88 },
        { ""Name"": ""Carol"", ""Score"": 92 }
    ]
}";
                File.WriteAllText(jsonFilePath, sampleJson);
            }

            // Read the JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook (empty Excel file)
            Workbook workbook = new Workbook();

            // Get the first worksheet where the data will be imported
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure JSON import options
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                // Treat JSON arrays as tables so that each element becomes a row
                ArrayAsTable = true,
                // Convert numeric strings and dates to proper Excel types
                ConvertNumericOrDate = true,
                // Optional: set number and date formats if needed
                NumberFormat = "#,##0.00",
                DateFormat = "yyyy-MM-dd"
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the workbook as an Excel file
            string excelOutputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.xlsx");
            workbook.Save(excelOutputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Excel file saved to: {excelOutputPath}");
        }
    }
}