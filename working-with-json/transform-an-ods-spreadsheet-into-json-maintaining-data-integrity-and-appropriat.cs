using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class OdsToJson
{
    static void Main()
    {
        // Load the ODS spreadsheet
        string odsPath = "input.ods";
        Workbook workbook = new Workbook(odsPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the used range of the worksheet
        Aspose.Cells.Range range = worksheet.Cells.MaxDisplayRange;

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(range, jsonOptions);

        // Write the JSON output to a file
        File.WriteAllText("output.json", json);
    }
}