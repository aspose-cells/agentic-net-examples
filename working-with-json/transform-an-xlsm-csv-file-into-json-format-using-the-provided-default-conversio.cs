using System;
using Aspose.Cells;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file (it may have .xlsm extension but contains CSV data)
            string sourcePath = "input.xlsm";

            // Path for the resulting JSON file
            string jsonPath = "output.json";

            // Load the CSV file using LoadOptions with CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Use default JSON save options (no custom settings)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {jsonPath}");
        }
    }
}