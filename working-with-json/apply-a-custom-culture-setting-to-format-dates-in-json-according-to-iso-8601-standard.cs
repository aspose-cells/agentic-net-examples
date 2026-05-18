using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Sample JSON containing date strings (any format, they will be parsed and reformatted)
        string jsonData = @"[
            { ""Id"": 1, ""Date"": ""15/05/2023 13:45:30"" },
            { ""Id"": 2, ""Date"": ""2023-06-20T09:15:00"" }
        ]";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure culture‑independent formatting (ISO 8601) by using InvariantCulture
        workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;

        // Configure JSON import options:
        // - Convert strings that look like dates to actual DateTime values
        // - Use ISO 8601 format when writing dates back to JSON
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,
            ConvertNumericOrDate = true,
            DateFormat = "yyyy-MM-ddTHH:mm:ss"   // ISO 8601 pattern
        };

        // Import the JSON data into the worksheet using the layout options
        JsonUtility.ImportData(jsonData, worksheet.Cells, 0, 0, layoutOptions);

        // Prepare JSON save options (optional indentation for readability)
        JsonSaveOptions saveOptions = new JsonSaveOptions
        {
            Indent = "  ", // two spaces
            ExportArea = new CellArea
            {
                StartRow = 0,
                EndRow = worksheet.Cells.MaxDataRow,
                StartColumn = 0,
                EndColumn = worksheet.Cells.MaxDataColumn
            },
            HasHeaderRow = true
        };

        // Save the workbook as JSON; dates will appear in ISO 8601 format
        workbook.Save("output.json", saveOptions);
    }
}