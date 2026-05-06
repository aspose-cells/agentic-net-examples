using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfToJson
{
    public class Converter
    {
        public static void ConvertExcelToJson(string excelPath, string jsonOutputPath)
        {
            // Ensure the source Excel file exists; if not, create a simple workbook for demonstration.
            if (!File.Exists(excelPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["A2"].PutValue("Value1");
                ws.Cells["B2"].PutValue("Value2");
                wb.Save(excelPath);
            }

            var workbook = new Workbook(excelPath);

            var jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            workbook.Save(jsonOutputPath, jsonOptions);
        }

        public static void Main()
        {
            string sourceExcel = Path.Combine(Directory.GetCurrentDirectory(), "ReportFromPdf.xlsx");
            string targetJson = Path.Combine(Directory.GetCurrentDirectory(), "ReportFromPdf.json");

            ConvertExcelToJson(sourceExcel, targetJson);
            Console.WriteLine("Conversion completed. JSON saved to: " + targetJson);
        }
    }
}