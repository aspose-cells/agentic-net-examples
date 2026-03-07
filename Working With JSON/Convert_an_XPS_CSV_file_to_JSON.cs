using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

class CsvToJsonConverter
{
    static void Main()
    {
        string csvPath = "input.csv";
        string jsonPath = "output.json";

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        cells.ImportCSV(csvPath, ",", true, 0, 0);

        int lastRow = cells.MaxDataRow;
        int lastColumn = cells.MaxDataColumn;

        AsposeRange dataRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            ExportEmptyCells = true
        };

        string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        File.WriteAllText(jsonPath, jsonOutput);

        Console.WriteLine("CSV file has been successfully converted to JSON.");
    }
}