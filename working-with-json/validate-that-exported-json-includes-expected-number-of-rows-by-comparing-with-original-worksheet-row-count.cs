using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonRowCountValidator
{
    static void Main()
    {
        // ---------- Create a workbook and populate data ----------
        Workbook workbook = new Workbook();                     // create workbook
        Worksheet sheet = workbook.Worksheets[0];              // get first worksheet
        Cells cells = sheet.Cells;

        // Fill some data with a few empty rows in between
        cells["A1"].PutValue("Header1");
        cells["B1"].PutValue("Header2");
        cells["A2"].PutValue("Row1Col1");
        cells["B2"].PutValue("Row1Col2");
        // Row 3 left empty intentionally
        cells["A4"].PutValue("Row3Col1");
        cells["B4"].PutValue("Row3Col2");

        // Determine the expected number of rows.
        // MaxDataRow returns the zero‑based index of the last row that contains data.
        // Adding 1 gives the total count of rows that have data (including empty rows before the last data row).
        int expectedRowCount = sheet.Cells.MaxDataRow + 1;

        // ---------- Export the worksheet to JSON ----------
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            // Ensure empty rows are NOT skipped so the JSON reflects the original row layout.
            SkipEmptyRows = false,
            // Export as a simple array (default) because we have only one sheet.
            AlwaysExportAsJsonObject = false
        };

        string jsonPath = "exported.json";
        workbook.Save(jsonPath, saveOptions);                  // save workbook as JSON

        // ---------- Load the JSON and count rows ----------
        string jsonContent = File.ReadAllText(jsonPath);

        // The default export format for a single sheet is a JSON array where each element represents a row.
        // Parse the JSON and count the array elements.
        using JsonDocument doc = JsonDocument.Parse(jsonContent);
        JsonElement root = doc.RootElement;

        int exportedRowCount = 0;

        if (root.ValueKind == JsonValueKind.Array)
        {
            exportedRowCount = root.GetArrayLength();
        }
        else if (root.ValueKind == JsonValueKind.Object)
        {
            // When AlwaysExportAsJsonObject is true, each sheet is a property.
            // Retrieve the first property (the sheet name) and count its array elements.
            foreach (JsonProperty prop in root.EnumerateObject())
            {
                if (prop.Value.ValueKind == JsonValueKind.Array)
                {
                    exportedRowCount = prop.Value.GetArrayLength();
                }
                break; // only need the first sheet for this validation
            }
        }

        // ---------- Validate ----------
        Console.WriteLine($"Expected row count (from worksheet): {expectedRowCount}");
        Console.WriteLine($"Exported row count (from JSON): {exportedRowCount}");

        if (expectedRowCount == exportedRowCount)
        {
            Console.WriteLine("Validation succeeded: row counts match.");
        }
        else
        {
            Console.WriteLine("Validation failed: row counts do not match.");
        }
    }
}