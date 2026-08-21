// Title: Export Excel Table to Indented JSON with Headers Using Aspose.Cells for .NET
// Description: Creates a workbook, adds a header row (Name, Age, City) and data rows, then uses JsonSaveOptions to treat the first row as keys, apply a four‑space indent, and export only the A1:C3 range to a formatted JSON file (TableExport.json).
// Keywords: Aspose.Cells | .NET | C# | Excel to JSON | JsonSaveOptions | header row as keys | JSON indentation | export specific range | pretty‑printed JSON | Aspose.Cells example
// Common Searches: Aspose.Cells export range to JSON C# | How to include column headers in JSON export with Aspose.Cells | Set indentation when saving workbook as JSON using Aspose.Cells | Export only selected cells to JSON file in .NET | JsonSaveOptions example for pretty printed JSON
// Developer Intent: Generate a readable JSON file from a defined Excel range, using the first row for property names and applying custom indentation.
// Use Cases: Convert a lookup table stored in Excel into a configuration JSON for an application. | Create an API response payload by exporting selected worksheet rows to JSON. | Produce a human‑friendly JSON report for data exchange between business systems.
// AI Prompts: Write C# code with Aspose.Cells that exports a worksheet range to a pretty‑printed JSON file, using the first row as keys. | Show how to change the indentation size and export area in JsonSaveOptions for a larger Excel table. | Explain how to include formula results or formatted values when exporting data to JSON with Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, adds a header row (Name, Age, City) and data rows, then uses JsonSaveOptions to treat the first row as keys, apply a four‑space indent, and export only the A1:C3 range to a formatted JSON file (TableExport.json).
class ExportTableToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("City");

        // Add data rows
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["C2"].PutValue("New York");

        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["C3"].PutValue("London");

        // Configure JSON save options: include headers, set indent, and define export area
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,                     // Use first row as keys
            Indent = "    ",                         // 4 spaces for readability
            ExportArea = new CellArea                // Export only the populated range (A1:C3)
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 0,
                EndColumn = 2
            }
        };

        // Save the workbook as a formatted JSON file
        workbook.Save("TableExport.json", saveOptions);
    }
}
