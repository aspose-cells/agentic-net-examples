using System;
using Aspose.Cells;

namespace AsposeCellsJsonConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Path where the JSON output will be saved
            string jsonPath = "output.json";

            // Load the existing workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create default JSON save options (no custom settings)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as a JSON file using the default options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
        }
    }
}