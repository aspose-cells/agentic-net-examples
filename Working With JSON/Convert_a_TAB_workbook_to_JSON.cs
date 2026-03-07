using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TabWorkbookToJson
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            string sourcePath = "input_tab_workbook.xlsx";

            // Create a sample workbook if the source file does not exist
            if (!File.Exists(sourcePath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Name = "Sheet1";

                // Add header row
                ws.Cells["A1"].PutValue("Id");
                ws.Cells["B1"].PutValue("Name");
                ws.Cells["C1"].PutValue("Value");

                // Add some data rows
                ws.Cells["A2"].PutValue(1);
                ws.Cells["B2"].PutValue("Item1");
                ws.Cells["C2"].PutValue(100);

                ws.Cells["A3"].PutValue(2);
                ws.Cells["B3"].PutValue("Item2");
                ws.Cells["C3"].PutValue(200);

                wb.Save(sourcePath);
            }

            Workbook workbook = new Workbook(sourcePath);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                ExportNestedStructure = false
            };

            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Workbook \"{sourcePath}\" has been converted to JSON and saved as \"{outputPath}\".");
        }
    }
}