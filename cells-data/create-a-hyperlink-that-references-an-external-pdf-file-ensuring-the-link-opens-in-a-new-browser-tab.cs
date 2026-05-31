using System;
using Aspose.Cells;

class CreatePdfHyperlink
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the display text for the hyperlink
        worksheet.Cells["A1"].PutValue("Open PDF Document");

        // Add a hyperlink to an external PDF file (opens in a new tab when saved as HTML)
        // Parameters: start cell, total rows, total columns, address of the PDF
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/document.pdf");

        // Configure HTML save options to use target="_blank" (new browser tab)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Save the workbook as an HTML file with the specified options
        workbook.Save("PdfHyperlink.html", htmlOptions);
    }
}