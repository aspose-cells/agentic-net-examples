using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class CsvToJsonConverter
{
    static void Main()
    {
        string csvPath = "input.csv";

        Workbook workbook = new Workbook();

        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        cells.ImportCSV(csvPath, ",", true, 0, 0);

        Aspose.Cells.Range usedRange = cells.MaxDisplayRange;

        ExportRangeToJsonOptions jsonOptions = new ExportRangeToJsonOptions
        {
            HasHeaderRow = true,
            ExportAsString = false
        };

        string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        string jsonPath = "output.json";
        File.WriteAllText(jsonPath, json);

        workbook.Save("intermediate.xlsx", SaveFormat.Xlsx);
    }
}