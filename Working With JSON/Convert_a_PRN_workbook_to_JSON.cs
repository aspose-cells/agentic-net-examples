using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPrnToJson
{
    class Program
    {
        static void Main()
        {
            // Path to the source PRN file
            string sourcePrnPath = "input.prn";

            Workbook workbook;

            if (File.Exists(sourcePrnPath))
            {
                // Load the workbook (format is auto‑detected)
                workbook = new Workbook(sourcePrnPath);
            }
            else
            {
                // Create a sample workbook if the PRN file is missing
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample");
                sheet.Cells["B1"].PutValue(123);
                sheet.Cells["A2"].PutValue(DateTime.Now);
            }

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,
                ExportEmptyCells = false
            };

            // Path for the resulting JSON file
            string outputJsonPath = "output.json";

            // Save the workbook as JSON using the configured options
            workbook.Save(outputJsonPath, jsonOptions);

            Console.WriteLine($"Workbook has been converted to JSON and saved at: {outputJsonPath}");
        }
    }
}