using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Configure JSON save options with pretty‑print indentation (4 spaces)
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                Indent = "    ",          // 4 spaces for indentation
                HasHeaderRow = true      // Treat first row as header
            };

            // Save the workbook as a formatted JSON file
            string outputPath = "FormattedOutput.json";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to JSON with indentation at: {outputPath}");
        }
    }
}