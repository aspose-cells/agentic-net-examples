using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row with camelCase property names
            sheet.Cells["A1"].PutValue("firstName");
            sheet.Cells["B1"].PutValue("lastName");
            sheet.Cells["C1"].PutValue("age");

            // Add sample data rows
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue("Doe");
            sheet.Cells["C2"].PutValue(30);

            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue("Smith");
            sheet.Cells["C3"].PutValue(25);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Ensure the output is a JSON object even if only one worksheet exists
                AlwaysExportAsJsonObject = true,
                // Treat the first row as header to use the camelCase column names as JSON keys
                HasHeaderRow = true,
                // Optional: format the JSON with indentation for readability
                Indent = "    ",
                // Export empty cells as null (adjust as needed)
                ExportEmptyCells = true,
                // Do not nest structures; flat array of objects
                ExportNestedStructure = false
            };

            // Define output path
            string outputPath = Path.Combine(Environment.CurrentDirectory, "output.json");

            // Save the workbook as JSON (lifecycle save rule)
            workbook.Save(outputPath, jsonOptions);

            // Display the generated JSON content
            Console.WriteLine("JSON exported to: " + outputPath);
            Console.WriteLine(File.ReadAllText(outputPath));
        }
    }
}