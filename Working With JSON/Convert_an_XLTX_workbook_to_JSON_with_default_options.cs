using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLTX template
            string sourcePath = "template.xltx";

            // Path for the output JSON file
            string jsonPath = "output.json";

            // Ensure the template exists; if not, create a simple one
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var sheet = tempWb.Worksheets[0];
                sheet.Name = "Sheet1";
                sheet.Cells["A1"].PutValue("Id");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                tempWb.Save(sourcePath, SaveFormat.Xltx);
            }

            // Load the XLTX workbook
            var workbook = new Workbook(sourcePath);

            // Create JSON save options (default)
            var jsonOptions = new JsonSaveOptions();

            // Save the workbook as JSON
            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Workbook '{sourcePath}' has been converted to JSON at '{jsonPath}'.");
        }
    }
}