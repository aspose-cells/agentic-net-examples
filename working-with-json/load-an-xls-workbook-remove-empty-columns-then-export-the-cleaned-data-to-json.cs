// Title: C# – Remove Blank Columns from an Excel Workbook and Export to JSON with Aspose.Cells
// Description: Load an XLSX file using Aspose.Cells for .NET, delete columns that are completely empty, configure JsonSaveOptions to skip empty rows, omit null cells, and treat the first row as a header, then save the cleaned worksheet as a compact JSON file.
// Keywords: Aspose.Cells C# remove blank columns | Excel to JSON .NET | JsonSaveOptions SkipEmptyRows | delete empty columns Aspose.Cells | export worksheet as JSON | clean Excel data before JSON conversion
// Common Searches: how to delete blank columns in Excel with Aspose.Cells | export cleaned Excel sheet to JSON C# | skip empty rows when saving Excel as JSON | Aspose.Cells remove empty columns before JSON export | C# convert Excel to JSON without empty columns
// Developer Intent: Load an Excel workbook, eliminate any columns that contain no data, and serialize the remaining content to a JSON file using Aspose.Cells for .NET.
// Use Cases: Prepare spreadsheet data for web APIs by stripping unused columns and generating lightweight JSON payloads. | Create front‑end data sources for JavaScript charts where only populated columns are needed. | Reduce storage and transmission size of Excel‑derived JSON by omitting blank columns and rows.
// AI Prompts: Show how to also delete blank rows after removing empty columns with Aspose.Cells. | Give a C# example that deserializes the exported JSON into strongly‑typed objects using Newtonsoft.Json. | Explain how to include cell formatting (e.g., number formats) in the JSON output with JsonSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Json;

// Load an XLSX file using Aspose.Cells for .NET, delete columns that are completely empty, configure JsonSaveOptions to skip empty rows, omit null cells, and treat the first row as a header, then save the cleaned worksheet as a compact JSON file.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove all columns that are completely blank
        worksheet.Cells.DeleteBlankColumns();

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Skip rows that are empty after column cleanup
            SkipEmptyRows = true,
            // Do not export empty cells as null (optional)
            ExportEmptyCells = false,
            // Treat the first row as header (optional, adjust as required)
            HasHeaderRow = true
        };

        // Path for the resulting JSON file
        string outputPath = "output.json";

        // Save the cleaned workbook as JSON using the configured options
        workbook.Save(outputPath, jsonOptions);

        Console.WriteLine("Workbook cleaned and exported to JSON successfully.");
    }
}
