using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the tab‑separated values (TSV) file
            string tsvPath = "input.tsv";

            // Path for the resulting JSON file
            string jsonPath = "output.json";

            // Load the TSV file into the workbook using appropriate load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.TabDelimited);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,
                ExportEmptyCells = true,
                HasHeaderRow = false,
                SkipEmptyRows = false,
                ExportNestedStructure = true
            };

            // Save the workbook as a JSON file
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"TSV data has been exported to JSON at: {jsonPath}");
        }
    }
}