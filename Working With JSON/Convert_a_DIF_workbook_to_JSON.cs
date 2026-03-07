using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Loading;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            DifToJsonConverter.Run();
        }
    }

    public class DifToJsonConverter
    {
        public static void Run()
        {
            // Path to the source DIF file
            string difPath = "input.dif";

            // Ensure the DIF file exists (create a simple one if not)
            if (!File.Exists(difPath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["A2"].PutValue(123);
                ws.Cells["B2"].PutValue(456);
                tempWb.Save(difPath, SaveFormat.Dif);
            }

            // Path where the resulting JSON file will be saved
            string jsonPath = "output.json";

            // Load the DIF workbook using DifLoadOptions
            var difLoadOptions = new DifLoadOptions();
            var workbook = new Workbook(difPath, difLoadOptions);

            // Configure JSON save options
            var jsonSaveOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,
                ExportEmptyCells = true,
                Indent = "  "
            };

            // Save the workbook as JSON
            workbook.Save(jsonPath, jsonSaveOptions);

            Console.WriteLine($"DIF workbook '{difPath}' has been converted to JSON at '{jsonPath}'.");
        }
    }
}