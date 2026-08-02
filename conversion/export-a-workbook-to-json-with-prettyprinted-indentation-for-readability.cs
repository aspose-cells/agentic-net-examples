// Title: Export a Workbook to Pretty‑Printed JSON with Indentation Using Aspose.Cells (C#)
// Description: The sample creates a new Workbook, populates a worksheet with a header row and data, sets JsonSaveOptions to use four‑space indentation and to keep column names, and writes the result as a formatted JSON file.
// Keywords: Aspose.Cells | C# | export workbook to JSON | formatted JSON output | JsonSaveOptions indentation | include header row | Excel to JSON conversion | pretty printed JSON | JSON save options | data serialization
// Common Searches: C# Aspose.Cells save workbook as indented JSON | How to use JsonSaveOptions for pretty‑print in Aspose.Cells | Export Excel sheet to formatted JSON with header row C# | Aspose.Cells JSON indentation example | Pretty‑print JSON from Excel using Aspose.Cells
// Developer Intent: Generate a readable JSON file from an Excel workbook while preserving column headers.
// Use Cases: Provide a clean JSON payload for a REST API that consumes spreadsheet data. | Create a human‑friendly JSON report for logging or documentation purposes. | Transfer Excel data to downstream services that require column names in the JSON structure. | Produce a formatted JSON file for version‑controlled data exchange.
// AI Prompts: Write C# code with Aspose.Cells that exports an existing Excel file to JSON using tab characters for indentation. | Show how to change JsonSaveOptions to output compact (no indentation) JSON while still including the header row. | Generate an example that reads a workbook from a stream and saves it as pretty‑printed JSON with a custom indent size. | Explain how to handle large worksheets when exporting to formatted JSON with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    // The sample creates a new Workbook, populates a worksheet with a header row and data, sets JsonSaveOptions to use four‑space indentation and to keep column names, and writes the result as a formatted JSON file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill sample data (including a header row)
            worksheet.Cells["A1"].PutValue("Name");
            worksheet.Cells["B1"].PutValue("Age");
            worksheet.Cells["A2"].PutValue("John");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["A3"].PutValue("Jane");
            worksheet.Cells["B3"].PutValue(25);

            // Set JSON save options with pretty‑printed indentation
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                Indent = "    ", // 4 spaces for each indent level
                HasHeaderRow = true
            };

            // Save the workbook as a formatted JSON file
            string outputPath = "pretty_output.json";
            workbook.Save(outputPath, saveOptions);
        }
    }
}
