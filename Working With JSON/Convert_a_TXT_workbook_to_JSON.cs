using System;
using Aspose.Cells;

class TxtToJsonConverter
{
    static void Main()
    {
        // Load the TXT (CSV) file into a workbook
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook("input.txt", loadOptions);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int maxRow = sheet.Cells.MaxDataRow;
        int maxColumn = sheet.Cells.MaxDataColumn;
        Aspose.Cells.Range dataRange = sheet.Cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,
            ExportEmptyCells = true
        };

        // Export the range to a JSON string (optional, for display)
        string jsonString = dataRange.ToJson(jsonOptions);
        Console.WriteLine(jsonString);

        // Save the entire workbook as a JSON file
        workbook.Save("output.json", jsonOptions);
    }
}