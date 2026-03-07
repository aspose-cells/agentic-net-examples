using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class DifToJsonConverter
    {
        public static void Run()
        {
            // Path to the source DIF file
            string sourcePath = "input.dif";

            // Path where the resulting JSON will be saved
            string jsonOutputPath = "output.json";

            // Load the DIF file using DifLoadOptions
            DifLoadOptions loadOptions = new DifLoadOptions();
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Get the first worksheet (or any worksheet you need)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to export – here we use the used range of the worksheet
            AsposeRange exportRange = worksheet.Cells.MaxDisplayRange;

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // Include empty cells in the output
                ExportEmptyCells = true,
                // Treat the first row as header (optional, set as needed)
                HasHeaderRow = true
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

            // Write the JSON string to the output file
            File.WriteAllText(jsonOutputPath, jsonResult);

            Console.WriteLine($"DIF file '{sourcePath}' has been converted to JSON and saved as '{jsonOutputPath}'.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            DifToJsonConverter.Run();
        }
    }
}