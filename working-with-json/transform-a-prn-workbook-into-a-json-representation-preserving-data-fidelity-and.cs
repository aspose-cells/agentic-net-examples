using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPrnToJson
{
    class Program
    {
        static void Main()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string sourcePath = Path.Combine(baseDir, "source.xlsx");

            if (!File.Exists(sourcePath))
            {
                Workbook sampleWb = new Workbook();
                Worksheet ws = sampleWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample");
                ws.Cells["B2"].PutValue(123);
                sampleWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            Workbook workbook = new Workbook(sourcePath);

            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            string jsonOutputPath = Path.Combine(baseDir, "output.json");
            workbook.Save(jsonOutputPath, jsonOptions);

            Console.WriteLine($"Workbook has been converted to JSON and saved to: {jsonOutputPath}");
        }
    }
}