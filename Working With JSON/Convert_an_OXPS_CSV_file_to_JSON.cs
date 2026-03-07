using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class OxpsCsvToJsonConverter
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source OXPS CSV file (treated as a CSV file)
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.oxps");

            // If the source file does not exist, create a simple CSV sample for demonstration
            if (!File.Exists(sourcePath))
            {
                string[] sampleLines =
                {
                    "Id,Name,Value",
                    "1,Alpha,100",
                    "2,Beta,200",
                    "3,Gamma,300"
                };
                File.WriteAllLines(sourcePath, sampleLines);
            }

            // Path for the resulting JSON file
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.json");

            // Load the CSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Configure JSON save options
            JsonSaveOptions jsonSaveOptions = new JsonSaveOptions
            {
                ExportNestedStructure = true,
                SkipEmptyRows = true,
                AlwaysExportAsJsonObject = true
            };

            // Save the workbook to JSON
            workbook.Save(outputPath, jsonSaveOptions);

            // Display the JSON content
            string jsonContent = File.ReadAllText(outputPath);
            Console.WriteLine("Conversion completed. JSON output:");
            Console.WriteLine(jsonContent);
        }
    }
}