using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        string csvFilePath = "input.txt";

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        cells.ImportCSV(csvFilePath, ",", true, 0, 0);

        var usedRange = cells.MaxDisplayRange;

        Aspose.Cells.Range exportRange = cells.CreateRange(
            usedRange.FirstRow,
            usedRange.FirstColumn,
            usedRange.RowCount,
            usedRange.ColumnCount);

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            ExportEmptyCells = true
        };

        string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, jsonOptions);

        File.WriteAllText("output.json", jsonOutput);

        Console.WriteLine("CSV data has been converted to JSON and saved to output.json");
    }
}