// Title: Export a full Aspose.Cells workbook to JSON with column headers (C#)
// Description: Shows how to create a workbook, add a header row, configure JsonSaveOptions (HasHeaderRow=true, ExportEmptyCells=false, ExportNestedStructure=false, Indent=" ") and save the entire workbook as a pretty‑printed JSON file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# JSON export | JsonSaveOptions | HasHeaderRow | ExportEmptyCells | ExportNestedStructure | pretty printed JSON | Excel to JSON .NET | full workbook export | sample code | Aspose.Cells JSON example
// Common Searches: Aspose.Cells export workbook to JSON C# | JsonSaveOptions HasHeaderRow example | Save Excel as indented JSON .NET | Export entire workbook as JSON Aspose | C# convert Excel to JSON with headers
// Developer Intent: Save the complete workbook as a JSON file while preserving the first row as column names.
// Use Cases: Generate a JSON payload for APIs directly from an Excel workbook that includes header information. | Create a human‑readable JSON configuration file from spreadsheet data for downstream processing. | Produce formatted JSON reports from Excel data for documentation, logging, or auditing.
// AI Prompts: Show how to modify JsonSaveOptions to include empty cells as null values in the exported JSON. | Provide code that exports only a selected worksheet to JSON while still including the header row. | Explain how to change the indentation style or remove it entirely when saving the workbook as JSON.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add a header row, configure JsonSaveOptions (HasHeaderRow=true, ExportEmptyCells=false, ExportNestedStructure=false, Indent=" ") and save the entire workbook as a pretty‑printed JSON file using Aspose.Cells for .NET.
class ExportWorkbookToJson
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Populate the first worksheet with a header row and some data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");   // Header
        sheet.Cells["B1"].PutValue("Age");    // Header
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Configure JsonSaveOptions to include the header row in the exported JSON
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,          // Include column headers
            ExportEmptyCells = false,     // Do not export empty cells as null
            ExportNestedStructure = false,
            Indent = "  "                 // Optional: pretty‑print with indentation
        };

        // Save the entire workbook as a JSON file using the options (lifecycle rule: save)
        workbook.Save("WorkbookExport.json", jsonOptions);
    }
}
