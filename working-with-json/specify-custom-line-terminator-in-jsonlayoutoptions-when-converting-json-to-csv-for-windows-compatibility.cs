using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class JsonToCsvWindowsLineTerminatorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // JSON data to be converted
                string jsonData = @"{
                    ""Employees"": [
                        { ""ID"": 1, ""Name"": ""John Doe"", ""Salary"": 50000 },
                        { ""ID"": 2, ""Name"": ""Jane Smith"", ""Salary"": 60000 }
                    ]
                }";

                // Configure JSON layout options
                JsonLayoutOptions layoutOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,          // Treat arrays as tables
                    ConvertNumericOrDate = true   // Auto‑convert numbers/dates
                };

                // Import JSON data into the worksheet
                JsonUtility.ImportData(jsonData, worksheet.Cells, 0, 0, layoutOptions);

                // Output CSV file path
                string outputPath = "Employees.csv";

                // Save the workbook as CSV (Windows line endings are used by default)
                workbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"CSV file saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            JsonToCsvWindowsLineTerminatorDemo.Run();
        }
    }
}