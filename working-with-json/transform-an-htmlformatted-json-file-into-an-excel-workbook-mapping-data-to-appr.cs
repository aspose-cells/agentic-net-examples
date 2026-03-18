using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonToExcel
{
    class Program
    {
        static void Main()
        {
            // Path to the JSON file (created if it does not exist)
            string jsonFilePath = Path.Combine(Path.GetTempPath(), "input.json");
            if (!File.Exists(jsonFilePath))
            {
                string sampleJson = @"{
                    ""Employees"": [
                        { ""Name"": ""John"", ""Age"": 30 },
                        { ""Name"": ""Jane"", ""Age"": 25 }
                    ]
                }";
                File.WriteAllText(jsonFilePath, sampleJson);
            }

            // Desired Excel output path
            string excelFilePath = Path.Combine(Environment.CurrentDirectory, "output.xlsx");
            string outputDir = Path.GetDirectoryName(excelFilePath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Configure JSON load options.
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                MultipleWorksheets = true
            };

            // Load the JSON file into a Workbook.
            Workbook workbook = new Workbook(jsonFilePath, loadOptions);

            // Example: rename first worksheet
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets[0].Name = "Employees";

            // Save the workbook as an Excel file.
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"JSON data has been converted and saved to '{excelFilePath}'.");
        }
    }
}