// Title: Export Excel Workbook to JSON with camelCase Property Names using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates sample rows, converts the header cells to camelCase, and saves the data as a formatted JSON file. The example demonstrates JsonSaveOptions settings such as AlwaysExportAsJsonObject, HasHeaderRow, Indent, ExportAsString = false, SkipEmptyRows, and an explicit ExportArea to control the output.
// Keywords: Aspose.Cells | C# | Excel to JSON | camelCase headers | JsonSaveOptions | AlwaysExportAsJsonObject | HasHeaderRow | ExportAsString false | SkipEmptyRows | export specific range | indent JSON output
// Common Searches: Aspose.Cells export Excel to JSON C# | how to convert Excel header row to camelCase with Aspose.Cells | JsonSaveOptions example with indentation and header row | skip empty rows when exporting Excel to JSON | export a selected range from workbook to JSON using Aspose.Cells
// Developer Intent: Export a workbook to JSON where each property name follows camelCase naming.
// Use Cases: Generate API payloads from spreadsheet data with camelCase field names while preserving numeric types. | Create configuration or settings files from template worksheets, exporting only a defined range with readable indentation. | Convert large Excel reports to JSON, automatically omitting blank rows and ensuring the result is a single JSON object.
// AI Prompts: Write C# code that uses Aspose.Cells to export a worksheet to JSON with camelCase column names and indented formatting. | Provide a helper method to transform Excel header values to camelCase before saving as JSON with JsonSaveOptions. | Explain how to configure JsonSaveOptions to always produce a JSON object, include a header row, skip empty rows, and export a specific cell area.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    // Creates a workbook, populates sample rows, converts the header cells to camelCase, and saves the data as a formatted JSON file. The example demonstrates JsonSaveOptions settings such as AlwaysExportAsJsonObject, HasHeaderRow, Indent, ExportAsString = false, SkipEmptyRows, and an explicit ExportArea to control the output.
    class Program
    {
        // Converts a string to camelCase (first character lower‑cased)
        static string ToCamelCase(string text)
        {
            if (string.IsNullOrEmpty(text) || char.IsLower(text[0]))
                return text;

            return char.ToLowerInvariant(text[0]) + text.Substring(1);
        }

        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (first row will be used as header)
            cells["A1"].PutValue("FirstName");
            cells["B1"].PutValue("LastName");
            cells["C1"].PutValue("Age");

            cells["A2"].PutValue("John");
            cells["B2"].PutValue("Doe");
            cells["C2"].PutValue(30);

            cells["A3"].PutValue("Jane");
            cells["B3"].PutValue("Smith");
            cells["C3"].PutValue(25);

            // Convert header row values to camelCase for JSON consistency
            int headerRow = 0; // zero‑based index
            for (int col = 0; col < 3; col++)
            {
                string originalHeader = cells[headerRow, col].StringValue;
                string camelHeader = ToCamelCase(originalHeader);
                cells[headerRow, col].PutValue(camelHeader);
            }

            // Configure JSON save options
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Export the range as a JSON object even if there is only one worksheet
                AlwaysExportAsJsonObject = true,
                // The first row contains header names
                HasHeaderRow = true,
                // Indent the output for readability (4 spaces)
                Indent = "    ",
                // Export values as their native types (not forced to string)
                ExportAsString = false,
                // Skip rows that contain no data
                SkipEmptyRows = true
            };

            // Define the export area (A1:C3) – optional, can be omitted to export the whole sheet
            saveOptions.ExportArea = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 2
            };

            // Save the workbook as a JSON file
            string outputPath = Path.Combine(Environment.CurrentDirectory, "output.json");
            workbook.Save(outputPath, saveOptions);

            // Display the generated JSON content
            Console.WriteLine("JSON exported to: " + outputPath);
            Console.WriteLine(File.ReadAllText(outputPath));
        }
    }
}
