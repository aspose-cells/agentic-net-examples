using System;
using Aspose.Cells;

class ExportWorksheetToHtml
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample text and a hyperlink
        worksheet.Cells["A1"].PutValue("Visit Aspose");
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Configure HTML save options so all links open in a new window/tab
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // equivalent to target="_blank"

        // Export the worksheet to HTML
        workbook.Save("output.html", saveOptions);
    }
}