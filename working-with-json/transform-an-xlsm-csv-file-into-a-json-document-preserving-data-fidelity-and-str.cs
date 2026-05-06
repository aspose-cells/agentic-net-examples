using System;
using Aspose.Cells;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file (it may have .xlsm extension)
            string sourceCsvPath = "input.xlsm";

            // Path where the resulting JSON will be saved
            string outputJsonPath = "output.json";

            // Load the CSV file into a workbook using CSV load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(sourceCsvPath, loadOptions);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // keep empty cells in the output
                HasHeaderRow = true,       // treat the first row as header
                ToExcelStruct = true       // preserve Excel‑like structure
            };

            // Save the workbook as JSON
            workbook.Save(outputJsonPath, jsonOptions);

            Console.WriteLine($"CSV data from '{sourceCsvPath}' has been converted to JSON and saved to '{outputJsonPath}'.");
        }
    }
}