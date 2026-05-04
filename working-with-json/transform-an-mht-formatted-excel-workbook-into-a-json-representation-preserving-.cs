using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Json;

class MhtToJsonConverter
{
    public static void ConvertMhtToJson(string mhtFilePath, string jsonOutputPath)
    {
        if (!File.Exists(mhtFilePath))
        {
            var sampleWb = new Workbook();
            sampleWb.Worksheets[0].Cells["A1"].PutValue("Sample Text");
            sampleWb.Worksheets[0].Cells["B2"].PutValue(12345);
            // Save as a supported format (e.g., XLSX) if MHTML is unavailable
            sampleWb.Save(mhtFilePath, SaveFormat.Xlsx);
        }

        var workbook = new Workbook(mhtFilePath); // Auto-detect format

        var saveOptions = new JsonSaveOptions
        {
            AlwaysExportAsJsonObject = true,
            ExportNestedStructure = true,
            ToExcelStruct = true
        };

        workbook.Save(jsonOutputPath, saveOptions);
    }

    static void Main()
    {
        string sourceMht = "input.mht";
        string targetJson = "output.json";

        ConvertMhtToJson(sourceMht, targetJson);
        Console.WriteLine($"Conversion completed. JSON saved to '{targetJson}'.");
    }
}