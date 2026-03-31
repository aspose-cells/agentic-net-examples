using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

class SpreadsheetMLToJson
{
    static void Main()
    {
        string inputPath = "input_spreadsheetml.xml";

        Workbook workbook;
        if (File.Exists(inputPath))
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.SpreadsheetML);
            workbook = new Workbook(inputPath, loadOptions);
        }
        else
        {
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Name = "Sheet1";
            ws.Cells["A1"].PutValue("Sample");
            ws.Cells["B2"].PutValue(123);
        }

        StringBuilder jsonWorkbook = new StringBuilder();
        jsonWorkbook.Append("{\"Worksheets\":[");

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange ?? sheet.Cells.CreateRange("A1");

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                ToExcelStruct = true
            };

            string sheetJson = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            jsonWorkbook.Append("{");
            jsonWorkbook.AppendFormat("\"Name\":\"{0}\",", sheet.Name);
            jsonWorkbook.AppendFormat("\"Data\":{0}", sheetJson);
            jsonWorkbook.Append("}");

            if (i < workbook.Worksheets.Count - 1)
                jsonWorkbook.Append(",");
        }

        jsonWorkbook.Append("]}");

        string outputPath = "workbook_structure.json";
        File.WriteAllText(outputPath, jsonWorkbook.ToString());

        Console.WriteLine($"Workbook has been converted to JSON and saved to '{outputPath}'.");
    }
}