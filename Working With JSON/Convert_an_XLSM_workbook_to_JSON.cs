using System;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    public class XlsmToJsonConverter
    {
        public static void Main(string[] args)
        {
            Convert();
        }

        public static void Convert()
        {
            // Path to the source XLSM workbook (contains macros)
            string sourcePath = "input.xlsm";

            // Desired output JSON file path
            string jsonPath = "output.json";

            // Load the XLSM workbook
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options (using default settings)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{jsonPath}'");
        }
    }
}