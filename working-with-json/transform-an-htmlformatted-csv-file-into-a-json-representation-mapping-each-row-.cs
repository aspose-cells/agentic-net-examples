using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class HtmlCsvToJson
{
    static void Main()
    {
        // Path to the HTML file that contains the CSV data as a table
        string htmlPath = "input.html";

        // Load the HTML file into a workbook with table conversion enabled
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        // Convert the first HTML table to a ListObject (optional but useful)
        loadOptions.TableLoadOptions.AddTableLoadOption(new HtmlTableLoadOption
        {
            TableIndex = 0,
            TableToListObject = true
        });

        // Create and load the workbook
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range (including header row)
        int lastRow = cells.MaxDataRow;
        int lastColumn = cells.MaxDataColumn;

        // Create a range that covers all populated cells
        Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

        // Configure JSON export options – assume first row contains headers
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            AlwaysExportAsJsonObject = true
        };

        // Export the range to a JSON string
        string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

        // Write the JSON string to an output file
        File.WriteAllText("output.json", json);
    }
}