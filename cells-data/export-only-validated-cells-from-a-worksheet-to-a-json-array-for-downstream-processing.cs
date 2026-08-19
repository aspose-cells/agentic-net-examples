// Title: Export Filled Cells from an Aspose.Cells Worksheet to JSON in C#
// Description: Creates a workbook, populates sample data, determines the used area, builds a range covering all occupied rows and columns, and uses JsonSaveOptions (ExportEmptyCells = false, HasHeaderRow = false, ExportAsString = false) to convert the range into a JSON array that contains only populated cells while preserving their native data types.
// Keywords: Aspose.Cells | .NET | C# | Excel to JSON | JsonSaveOptions | ExportEmptyCells false | skip empty cells | range.ToJson | used range export | worksheet serialization
// Common Searches: Aspose.Cells export only filled cells to JSON | C# convert Excel used range to JSON | JsonSaveOptions skip empty cells example | How to serialize worksheet range as JSON with Aspose.Cells | Export Excel data without blanks using Aspose.Cells .NET
// Developer Intent: Generate a JSON array that includes only the non‑empty cells of a worksheet for downstream processing.
// Use Cases: Send compact JSON payloads from Excel data to REST APIs, omitting blank values. | Create lightweight data extracts for reporting dashboards that preserve original numeric and string types. | Provide front‑end JavaScript applications with clean JSON representations of Excel tables without unnecessary empty entries.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet's used range to JSON while excluding empty cells and keeping original data types. | Show how to configure JsonSaveOptions to skip blank cells and treat the first row as data when converting a range to JSON. | Explain how to modify the sample to include a header row in the JSON output using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Creates a workbook, populates sample data, determines the used area, builds a range covering all occupied rows and columns, and uses JsonSaveOptions (ExportEmptyCells = false, HasHeaderRow = false, ExportAsString = false) to convert the range into a JSON array that contains only populated cells while preserving their native data types.
class ExportValidatedCellsToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including some empty cells)
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("John");
            // B2 intentionally left empty
            cells["A3"].PutValue("Alice");
            cells["B3"].PutValue(25);

            // Determine the used area of the worksheet
            int maxRow = cells.MaxDataRow;      // zero‑based index of last row with data
            int maxCol = cells.MaxDataColumn;   // zero‑based index of last column with data

            // Create a range that covers all used cells
            Aspose.Cells.Range range = cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);

            // Set JSON export options:
            // - ExportEmptyCells = false  => skip empty cells (only "validated"/filled cells are exported)
            // - HasHeaderRow = false      => treat first row as data, not as header
            // - ExportAsString = false    => keep original data types
            JsonSaveOptions options = new JsonSaveOptions
            {
                ExportEmptyCells = false,
                HasHeaderRow = false,
                ExportAsString = false
            };

            // Convert the range to a JSON string using the configured options
            string json = range.ToJson(options);

            // Output the resulting JSON array
            Console.WriteLine(json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
