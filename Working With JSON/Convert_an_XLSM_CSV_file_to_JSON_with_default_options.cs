using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file (it may have .xlsm extension)
            string sourcePath = "input.xlsm";   // replace with actual file path
            // Desired JSON output file path
            string jsonPath = "output.json";

            // Load the CSV file using LoadOptions with CSV format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook as JSON using default JsonSaveOptions
            JsonSaveOptions jsonOptions = new JsonSaveOptions(); // default options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Conversion completed. JSON saved to: {jsonPath}");
        }
    }
}