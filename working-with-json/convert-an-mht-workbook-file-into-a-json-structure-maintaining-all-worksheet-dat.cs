using System;
using Aspose.Cells;

class MhtToJsonConverter
{
    static void Main()
    {
        string sourceMhtPath = "input.mht";

        Workbook workbook;
        if (System.IO.File.Exists(sourceMhtPath))
        {
            workbook = new Workbook(sourceMhtPath);
        }
        else
        {
            workbook = new Workbook();
            var ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample");
            ws.Cells["B2"].PutValue(123);
        }

        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportStylePool = true,
            ExportEmptyCells = true,
            ExportNestedStructure = true
        };

        string outputJsonPath = "output.json";
        workbook.Save(outputJsonPath, jsonOptions);

        Console.WriteLine($"MHTML workbook has been converted to JSON and saved to '{outputJsonPath}'.");
    }
}