using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class MHtmlCsvToJson
{
    static void Main()
    {
        // Input MHTML file that contains CSV data
        string mhtmlPath = "input.mhtml";

        // If the MHTML file does not exist, create a sample workbook and save it as MHTML
        if (!File.Exists(mhtmlPath))
        {
            Workbook sampleWb = new Workbook();
            Worksheet ws = sampleWb.Worksheets[0];
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Age");
            ws.Cells["A2"].PutValue("Alice");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("Bob");
            ws.Cells["B3"].PutValue(25);
            sampleWb.Save(mhtmlPath, SaveFormat.MHtml);
        }

        // Output JSON file path
        string jsonPath = "output.json";

        // Load the MHTML file as a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.MHtml);
        Workbook workbook = new Workbook(mhtmlPath, loadOptions);

        // Access the first worksheet (CSV data is expected here)
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range (including header row)
        int lastRow = worksheet.Cells.MaxDataRow;
        int lastColumn = worksheet.Cells.MaxDataColumn;

        // Create a range that covers all used cells
        Aspose.Cells.Range dataRange = worksheet.Cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,          // First row contains column names
            ExportEmptyCells = true,      // Preserve empty cells
            ExportNestedStructure = false
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"MHTML CSV successfully converted to JSON at: {jsonPath}");
    }
}