using System;
using System.IO;
using Aspose.Cells;

class CsvToJsonConverter
{
    static void Main()
    {
        string csvPath = "input.csv";
        string jsonPath = "output.json";

        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        cells.ImportCSV(csvPath, ",", true, 0, 0);

        int maxRow = cells.MaxDataRow;
        int maxColumn = cells.MaxDataColumn;

        Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            AlwaysExportAsJsonObject = true,
            ToExcelStruct = true
        };

        string json = dataRange.ToJson(jsonOptions);

        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"CSV file '{csvPath}' has been converted to JSON and saved as '{jsonPath}'.");
    }
}