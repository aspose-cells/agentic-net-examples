using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOxpsToJson
{
    class Program
    {
        static void Main()
        {
            // Determine the full path of the input OXPS file (located in the executable directory)
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.oxps");
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.json");

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Input file not found: {sourcePath}");
                return;
            }

            // Load the workbook (auto-detect format)
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = true
            };

            // Save as JSON
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"OXPS workbook has been successfully converted to JSON at '{outputPath}'.");
        }
    }
}