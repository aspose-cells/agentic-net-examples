using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

class MhtmlToJsonConverter
{
    static void Main()
    {
        string mhtmlPath = "input.mht";

        Workbook workbook;
        if (File.Exists(mhtmlPath))
        {
            workbook = new Workbook(mhtmlPath);
        }
        else
        {
            // Create a sample workbook if the MHTML file is missing
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Name = "SampleSheet";
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Age");
            ws.Cells["A2"].PutValue("John");
            ws.Cells["B2"].PutValue(30);
        }

        var jsonBuilder = new StringBuilder();
        jsonBuilder.Append("{");

        bool firstSheet = true;

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int lastRow = sheet.Cells.MaxDataRow;
            int lastCol = sheet.Cells.MaxDataColumn;

            if (lastRow < 0 || lastCol < 0)
                continue;

            var usedRange = sheet.Cells.CreateRange(0, 0, lastRow + 1, lastCol + 1);

            var jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true
            };

            string sheetJson = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

            if (!firstSheet)
                jsonBuilder.Append(",");

            jsonBuilder.Append($"\"{sheet.Name}\": {sheetJson}");
            firstSheet = false;
        }

        jsonBuilder.Append("}");

        Console.WriteLine(jsonBuilder.ToString());
    }
}