using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        string csvPath = "input.csv";
        string jsonPath = "output.json";

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        worksheet.Cells.ImportCSV(csvPath, ",", true, 0, 0);

        int lastRow = worksheet.Cells.MaxDataRow;
        int lastCol = worksheet.Cells.MaxDataColumn;

        Aspose.Cells.Range dataRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastCol + 1);

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true
        };

        string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"CSV has been converted to JSON and saved to '{jsonPath}'.");
    }
}