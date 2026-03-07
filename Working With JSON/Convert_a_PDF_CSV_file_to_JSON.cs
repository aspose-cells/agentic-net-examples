using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        string sourcePath = "input.csv";

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        worksheet.Cells.ImportCSV(sourcePath, ",", true, 0, 0);

        int lastRow = worksheet.Cells.MaxDataRow;
        int lastColumn = worksheet.Cells.MaxDataColumn;

        Aspose.Cells.Range dataRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,
            ExportNestedStructure = false
        };

        string jsonOutput = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        File.WriteAllText("output.json", jsonOutput);

        Console.WriteLine("CSV data has been successfully converted to JSON.");
    }
}