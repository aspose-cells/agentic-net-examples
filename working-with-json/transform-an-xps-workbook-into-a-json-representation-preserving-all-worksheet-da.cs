using System;
using System.IO;
using Aspose.Cells;

namespace XpsToJsonConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Ensure the workbook exists; if not, create a sample workbook and save it as XLSX
            if (!File.Exists(sourcePath))
            {
                // Create a sample workbook
                Workbook sampleWb = new Workbook();
                Worksheet sheet = sampleWb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["A2"].PutValue("Alice");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Bob");
                sheet.Cells["B3"].PutValue(25);

                // Save the sample workbook as XLSX
                sampleWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Configure JSON save options to preserve formatting details
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportStylePool = true,
                ExportEmptyCells = true,
                MergeAreas = true,
                ToExcelStruct = false
            };

            // Save the workbook as a JSON file with the specified options
            workbook.Save("output.json", jsonOptions);
        }
    }
}