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
            string csvPath = "data.tsv";
            string jsonPath = "data.json";

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells.ImportCSV(csvPath, "\t", true, 0, 0);

            Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,
                ExportEmptyCells = true,
                Indent = "    "
            };

            string jsonOutput = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            File.WriteAllText(jsonPath, jsonOutput);

            Console.WriteLine($"CSV file \"{csvPath}\" has been converted to JSON and saved as \"{jsonPath}\".");
        }
    }
}