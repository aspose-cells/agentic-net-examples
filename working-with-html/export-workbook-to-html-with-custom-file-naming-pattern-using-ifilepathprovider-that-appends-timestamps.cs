using System;
using Aspose.Cells;

// Custom file path provider that appends a timestamp to each worksheet file name
class TimestampFilePathProvider : IFilePathProvider
{
    public string GetFullName(string sheetName)
    {
        // Create a timestamp string (e.g., 20231127_154530)
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        // Return the full file name for the worksheet HTML file
        return $"{sheetName}_{timestamp}.html";
    }
}

class ExportWorkbookToHtml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate the first worksheet
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "DataSheet";
        ws1.Cells["A1"].PutValue("Hello");
        ws1.Cells["A2"].PutValue("World");

        // Add a second worksheet to demonstrate separate HTML files per sheet
        Worksheet ws2 = workbook.Worksheets.Add("Summary");
        ws2.Cells["A1"].PutValue("Summary Info");

        // Set up HTML save options and assign the custom file path provider
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.FilePathProvider = new TimestampFilePathProvider();

        // Save the workbook as HTML; the main file name can be any name you choose
        workbook.Save("Workbook.html", saveOptions);
    }
}