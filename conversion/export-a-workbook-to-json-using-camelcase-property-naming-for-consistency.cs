using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        // Simple helper to convert a string to camelCase (first character lower‑cased)
        static string ToCamelCase(string text)
        {
            if (string.IsNullOrEmpty(text) || char.IsLower(text[0]))
                return text;

            return char.ToLowerInvariant(text[0]) + text.Substring(1);
        }

        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate data with a header row
            // Header values will be converted to camelCase to ensure consistency in the exported JSON
            string[] headers = { "FirstName", "LastName", "Age", "DateOfBirth" };
            for (int col = 0; col < headers.Length; col++)
            {
                // Convert each header to camelCase before putting it into the cell
                sheet.Cells[0, col].PutValue(ToCamelCase(headers[col]));
            }

            // Sample data rows
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue("Doe");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["D2"].PutValue(new DateTime(1993, 5, 12));

            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue("Smith");
            sheet.Cells["C3"].PutValue(28);
            sheet.Cells["D3"].PutValue(new DateTime(1995, 8, 23));

            // 3. Configure JSON save options
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Export as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // Treat the first row as header (already camelCase)
                HasHeaderRow = true,
                // Skip rows that contain no data
                SkipEmptyRows = true,
                // Optional: pretty‑print with 4‑space indentation
                Indent = "    "
            };

            // 4. Save the workbook as JSON
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ExportedData.json");
            workbook.Save(outputPath, saveOptions);

            // 5. Display the resulting JSON content
            string jsonContent = File.ReadAllText(outputPath);
            Console.WriteLine("Exported JSON:");
            Console.WriteLine(jsonContent);
        }
    }
}