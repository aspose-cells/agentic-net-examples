using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXpsToJson
{
    class Program
    {
        static void Main()
        {
            // Paths for the source XLSX file and final JSON output
            string sourceXlsxPath = "input.xlsx";
            string outputJsonPath = "output.json";

            // Ensure the source XLSX file exists; create a dummy one if it does not
            if (!File.Exists(sourceXlsxPath))
            {
                Workbook dummyWb = new Workbook();
                dummyWb.Worksheets[0].Cells["A1"].PutValue("Sample");
                dummyWb.Save(sourceXlsxPath, SaveFormat.Xlsx);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourceXlsxPath);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ToExcelStruct = true,
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportAsString = true,
                Indent = "  "
            };

            // Save the workbook as JSON
            workbook.Save(outputJsonPath, jsonOptions);

            Console.WriteLine($"Workbook has been converted to JSON and saved at: {outputJsonPath}");
        }
    }
}