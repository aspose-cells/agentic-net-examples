using System;
using System.IO;
using Aspose.Cells;

namespace OxpsToJsonExample
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "input.oxps";
            string outputPath = "output.json";

            if (!File.Exists(sourcePath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample");
                tempWb.Save(sourcePath);
            }

            Workbook workbook = new Workbook(sourcePath);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true
            };

            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"OXPS workbook has been converted to JSON and saved to '{outputPath}'.");
        }
    }
}