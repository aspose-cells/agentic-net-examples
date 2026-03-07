using System;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    public class DbfToJsonConverter
    {
        public static void Run()
        {
            // Path to the source DBF file
            string sourcePath = "input.dbf";

            // Path where the resulting JSON file will be saved
            string jsonPath = "output.json";

            // Load the DBF file using DbfLoadOptions
            DbfLoadOptions loadOptions = new DbfLoadOptions();
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export the workbook as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,

                // Include empty cells in the JSON output
                ExportEmptyCells = true,

                // Treat the first row as header names
                HasHeaderRow = true
            };

            // Save the workbook as JSON using the configured options
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"DBF file '{sourcePath}' has been successfully converted to JSON at '{jsonPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DbfToJsonConverter.Run();
        }
    }
}