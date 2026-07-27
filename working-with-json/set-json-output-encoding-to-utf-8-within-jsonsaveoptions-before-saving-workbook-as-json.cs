using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsJsonUtf8Example
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Create JSON save options (no Encoding property exists for JsonSaveOptions)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Example of other options you might set
                Indent = "  ",
                HasHeaderRow = true,
                ExportEmptyCells = false
            };

            // Save the workbook to a memory stream using the JSON options
            using (MemoryStream jsonStream = new MemoryStream())
            {
                workbook.Save(jsonStream, jsonOptions);

                // Ensure the stream is positioned at the beginning
                jsonStream.Position = 0;

                // Convert the stream bytes to a UTF‑8 encoded string
                string jsonContent = Encoding.UTF8.GetString(jsonStream.ToArray());

                // Write the UTF‑8 string to a file explicitly using UTF‑8 encoding
                string outputPath = "WorkbookOutput.json";
                File.WriteAllText(outputPath, jsonContent, Encoding.UTF8);

                Console.WriteLine($"Workbook saved as JSON with UTF‑8 encoding to: {outputPath}");
            }
        }
    }
}