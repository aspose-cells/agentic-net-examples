using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        string sourcePath = "input.sxc";

        Workbook workbook = new Workbook(sourcePath);
        Worksheet worksheet = workbook.Worksheets[0];

        int lastRow = worksheet.Cells.MaxDataRow;
        int lastColumn = worksheet.Cells.MaxDataColumn;

        string jsonResult;

        if (lastRow < 0 || lastColumn < 0)
        {
            jsonResult = "[]";
        }
        else
        {
            Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true
            };

            jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);
        }

        Console.WriteLine(jsonResult);
        File.WriteAllText("output.json", jsonResult);
    }
}