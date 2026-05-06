using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class OdsOrCsvToJsonConverter
{
    static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: OdsOrCsvToJsonConverter <inputFile> <outputJson>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        Workbook workbook = new Workbook(inputPath);
        Worksheet worksheet = workbook.Worksheets[0];

        int lastRow = worksheet.Cells.MaxDataRow;
        int lastColumn = worksheet.Cells.MaxDataColumn;

        AsposeRange dataRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            ExportEmptyCells = true,
            ExportNestedStructure = false
        };

        string jsonContent = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);
        File.WriteAllText(outputPath, jsonContent);

        Console.WriteLine($"Conversion completed. JSON saved to: {outputPath}");
    }
}