using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "input.xls";
            Workbook workbook = new Workbook(sourcePath);
            Worksheet worksheet = workbook.Worksheets[0];
            AsposeRange dataRange = worksheet.Cells.MaxDisplayRange;
            JsonSaveOptions jsonOptions = new JsonSaveOptions();
            string jsonResult = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);
            Console.WriteLine(jsonResult);
        }
    }
}