using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class HtmlToJsonConverter
{
    static void Main()
    {
        // Path to the HTML-formatted spreadsheet
        string htmlPath = "input.html";

        // Load the HTML file into a workbook using LoadOptions for HTML format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // Access the first worksheet (or iterate worksheets as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int totalRows = worksheet.Cells.MaxDataRow + 1;      // MaxDataRow is zero‑based
        int totalCols = worksheet.Cells.MaxDataColumn + 1;   // MaxDataColumn is zero‑based

        // Create a range that covers the entire used area
        Aspose.Cells.Range usedRange = worksheet.Cells.CreateRange(0, 0, totalRows, totalCols);

        // Configure JSON export options to preserve Excel structure
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ToExcelStruct = true   // Export as Excel‑style JSON structure
        };

        // Export the range to a JSON string
        string jsonResult = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        // Write the JSON output to a file
        File.WriteAllText("output.json", jsonResult);

        Console.WriteLine("HTML spreadsheet has been converted to JSON successfully.");
    }
}