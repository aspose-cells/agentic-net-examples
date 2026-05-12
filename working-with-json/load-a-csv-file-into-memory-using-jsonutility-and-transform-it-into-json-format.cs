using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCsvToJson
{
    class Program
    {
        static void Main()
        {
            string csvPath = "input.csv";

            File.WriteAllText(csvPath,
                "Name,Age,City\n" +
                "John,30,New York\n" +
                "Alice,25,London\n" +
                "Bob,35,Paris");

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            cells.ImportCSV(csvPath, ",", true, 0, 0);

            AsposeRange dataRange = cells.MaxDisplayRange;

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };

            string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            Console.WriteLine("JSON representation of the CSV data:");
            Console.WriteLine(jsonOutput);

            if (File.Exists(csvPath))
            {
                File.Delete(csvPath);
            }
        }
    }
}