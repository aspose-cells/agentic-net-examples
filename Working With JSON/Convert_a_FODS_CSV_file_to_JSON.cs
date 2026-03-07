using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class FodsToJsonConverter
    {
        public static void Run()
        {
            // Paths for source FODS file and destination JSON file
            string sourcePath = "input.fods";
            string jsonPath = "output.json";

            // Load the FODS file with appropriate load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

            // Configure JSON export options (customize as required)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,   // include empty cells in JSON
                HasHeaderRow = true,       // treat first row as header
                ExportNestedStructure = false
            };

            // Export the range to a JSON string
            string jsonContent = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            // Write the JSON string to the output file
            File.WriteAllText(jsonPath, jsonContent);

            Console.WriteLine($"Conversion completed. JSON saved to: {jsonPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            FodsToJsonConverter.Run();
        }
    }
}