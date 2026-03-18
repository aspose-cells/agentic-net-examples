using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Loading;
using Aspose.Cells.Utility;

class DifToJsonConverter
{
    static void Main()
    {
        string difPath = "input.dif";
        DifLoadOptions loadOptions = new DifLoadOptions();
        Workbook workbook = new Workbook(difPath, loadOptions);

        StringBuilder jsonBuilder = new StringBuilder();
        jsonBuilder.Append("{");

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true
            };

            string sheetJson = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            jsonBuilder.Append($"\"{sheet.Name}\":{sheetJson}");

            if (i < workbook.Worksheets.Count - 1)
                jsonBuilder.Append(",");
        }

        jsonBuilder.Append("}");

        string outputPath = "output.json";
        File.WriteAllText(outputPath, jsonBuilder.ToString());

        Console.WriteLine($"DIF file has been converted to JSON and saved to '{outputPath}'.");
    }
}