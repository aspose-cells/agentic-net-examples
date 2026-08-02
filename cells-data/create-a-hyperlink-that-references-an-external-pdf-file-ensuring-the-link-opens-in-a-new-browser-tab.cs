using System;
using Aspose.Cells;

class CreatePdfHyperlink
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set the display text for the hyperlink
        sheet.Cells["A1"].PutValue("Open PDF");

        // Add a hyperlink to an external PDF file (opens when clicked)
        // Parameters: start cell, rows, columns, address of the PDF
        sheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/document.pdf");

        // Configure HTML save options so that links open in a new browser tab/window
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // _blank target

        // Save the workbook as an HTML file with the specified options
        workbook.Save("PdfLink.html", saveOptions);
    }
}