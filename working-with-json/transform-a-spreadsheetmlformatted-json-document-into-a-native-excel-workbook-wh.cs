using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the SpreadsheetML‑formatted JSON document
        string jsonFilePath = "input.json";

        // Path where the native Excel workbook will be saved
        string excelFilePath = "output.xlsx";

        // Configure JSON load options (adjust if the JSON contains multiple worksheets)
        JsonLoadOptions loadOptions = new JsonLoadOptions
        {
            MultipleWorksheets = false
        };

        // Load the JSON document into a Workbook instance
        Workbook workbook = new Workbook(jsonFilePath, loadOptions);

        // Save the workbook as a native Excel file
        workbook.Save(excelFilePath);
    }
}