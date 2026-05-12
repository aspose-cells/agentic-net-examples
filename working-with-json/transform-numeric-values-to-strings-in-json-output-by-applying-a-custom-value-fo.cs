using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsJsonStringExportDemo
{
    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Laptop");
            cells["B2"].PutValue(999.99);
            cells["A3"].PutValue("Phone");
            cells["B3"].PutValue(599.99);

            Style priceStyle = workbook.CreateStyle();
            priceStyle.Number = 2;
            priceStyle.Custom = "$#,##0.00";

            AsposeRange priceRange = cells.CreateRange("B2:B3");
            StyleFlag flag = new StyleFlag { All = true };
            priceRange.ApplyStyle(priceStyle, flag);

            ExportRangeToJsonOptions exportOptions = new ExportRangeToJsonOptions
            {
                ExportAsString = true,
                Indent = "    ",
                HasHeaderRow = true
            };

            AsposeRange exportRange = cells.CreateRange("A1:B3");
            string jsonOutput = JsonUtility.ExportRangeToJson(exportRange, exportOptions);

            Console.WriteLine("Exported JSON with numeric values as strings:");
            Console.WriteLine(jsonOutput);
        }
    }
}