using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsvDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Sample JSON data
                string jsonData = @"{
                    ""Employees"": [
                        { ""ID"": 1, ""Name"": ""John Doe"", ""Salary"": 50000 },
                        { ""ID"": 2, ""Name"": ""Jane Smith"", ""Salary"": 60000 }
                    ]
                }";

                // Create a new workbook (in‑memory)
                Workbook workbook = new Workbook();

                // Configure JSON layout options (treat arrays as tables, convert numbers/dates)
                JsonLayoutOptions jsonOptions = new JsonLayoutOptions
                {
                    ArrayAsTable = true,
                    ConvertNumericOrDate = true
                };

                // Import JSON data into the first worksheet starting at cell A1 (row 0, column 0)
                JsonUtility.ImportData(jsonData, workbook.Worksheets[0].Cells, 0, 0, jsonOptions);

                // Set up CSV save options with semicolon as the delimiter
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ';' // Define semicolon as CSV delimiter
                };

                // Define output path and ensure the directory exists
                string outputPath = "EmployeesOutput.csv";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a CSV file using the defined options
                workbook.Save(outputPath, csvOptions);
                Console.WriteLine($"CSV file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}