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

        Aspose.Cells.Range usedRange = worksheet.Cells.MaxDisplayRange;

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ToExcelStruct = true,
            AlwaysExportAsJsonObject = true
        };

        string jsonContent = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        File.WriteAllText(jsonPath, jsonContent);

        Console.WriteLine("CSV file has been successfully converted to JSON.");
    }
}