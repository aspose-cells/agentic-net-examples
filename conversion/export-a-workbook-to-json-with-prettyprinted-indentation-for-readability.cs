// Title: Export a Workbook to Pretty‑Printed JSON with Indentation using Aspose.Cells for C#
// Description: This example creates a new Workbook, adds sample names and ages to cells A1:B3, configures JsonSaveOptions with a 4‑space indent, includes a header row, limits the export to the defined range, and saves the result as FormattedOutput.json, producing a readable, indented JSON representation of the data.
// Keywords: Aspose.Cells | C# | .NET | JsonSaveOptions | pretty print JSON | JSON indentation | export workbook to JSON | formatted JSON output | cell range export | header row JSON
// Common Searches: Aspose.Cells export workbook to JSON C# | How to pretty print JSON with Aspose.Cells | JsonSaveOptions Indent property example | Export specific cell range to JSON Aspose.Cells | Include header row when exporting to JSON Aspose
// Developer Intent: Save spreadsheet data as a human‑readable, indented JSON file.
// Use Cases: Generate a clean JSON report from tabular data for API consumption or documentation. | Create version‑control‑friendly JSON files from defined worksheet ranges with column headers. | Export configuration tables from Excel to formatted JSON for easy editing and review.
// AI Prompts: Show how to change the JsonSaveOptions indent to a tab character instead of spaces. | Provide code that exports each worksheet in a workbook to separate pretty‑printed JSON files. | Explain how to include cell formatting (e.g., number formats, dates) in the JSON output using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    // This example creates a new Workbook, adds sample names and ages to cells A1:B3, configures JsonSaveOptions with a 4‑space indent, includes a header row, limits the export to the defined range, and saves the result as FormattedOutput.json, producing a readable, indented JSON representation of the data.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Configure JSON save options with pretty‑printed indentation (4 spaces)
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                Indent = "    ",               // 4 spaces for indentation
                HasHeaderRow = true,          // First row contains column names
                ExportArea = new CellArea     // Export the range A1:B3
                {
                    StartRow = 0,
                    EndRow = 2,
                    StartColumn = 0,
                    EndColumn = 1
                }
            };

            // Save the workbook as a formatted JSON file
            string outputPath = "FormattedOutput.json";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to JSON with indentation at: {outputPath}");
        }
    }
}
