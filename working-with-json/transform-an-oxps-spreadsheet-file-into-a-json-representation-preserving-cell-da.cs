using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class OxpsToJsonConverter
    {
        public static void Run()
        {
            // Path to the source OXPS spreadsheet file
            string sourcePath = "input.oxps";

            // Path where the resulting JSON file will be saved
            string jsonPath = "output.json";

            // Load the OXPS file into a Workbook object
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options to preserve the Excel structure
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON structure that mirrors the Excel layout
                ToExcelStruct = true,

                // Include empty cells in the output (optional, but helps preserve exact layout)
                ExportEmptyCells = true,

                // Export cell values as strings to retain original formatting (optional)
                ExportAsString = true
            };

            // Save the workbook as a JSON file using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"OXPS file has been converted to JSON and saved to '{jsonPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OxpsToJsonConverter.Run();
        }
    }
}