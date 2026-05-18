using System;
using Aspose.Cells;
using Aspose.Cells.Json;

class JsonEncodingExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Configure JSON save options.
        // JsonSaveOptions does not expose an Encoding property because
        // the JSON output is always written using UTF‑8 encoding internally.
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Optional: format the JSON with indentation for readability
            Indent = "    ",
            // Export the whole sheet (no specific area needed here)
            ExportArea = new CellArea { StartRow = 0, EndRow = 2, StartColumn = 0, EndColumn = 1 },
            HasHeaderRow = true
        };

        // Save the workbook as a JSON file. The resulting file is encoded in UTF‑8.
        workbook.Save("output.json", jsonOptions);

        Console.WriteLine("Workbook saved as UTF‑8 encoded JSON to 'output.json'.");
    }
}