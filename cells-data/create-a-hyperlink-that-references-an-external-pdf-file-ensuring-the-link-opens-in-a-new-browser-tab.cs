// Title: Generate an HTML file with a cell hyperlink to an external PDF that opens in a new browser tab using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to insert a hyperlink to a PDF URL in cell A1 and export the workbook to HTML with the link set to open in a new tab. | Show how to set HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank so that external links open in a new window when saving a workbook as HTML. | Demonstrate adding a worksheet cell hyperlink and saving the workbook as an HTML file while preserving the hyperlink's target behavior.
// Common Searches: Aspose.Cells C# add hyperlink to PDF and export as HTML with _blank target | How to make hyperlinks open in a new tab when saving Excel to HTML using Aspose.Cells | Set HtmlLinkTargetType to Blank for external links in Aspose.Cells HTML export | Create HTML link to external PDF from Excel cell using Aspose.Cells .NET
// Tags: Aspose.Cells PDF hyperlink cell | HtmlSaveOptions target blank | HTML export external links Aspose.Cells | C# set hyperlink new tab | save workbook as HTML Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, places a hyperlink in cell A1 that points to an external PDF, configures HtmlSaveOptions.LinkTargetType to Blank so the link opens in a new browser tab, and saves the workbook as an HTML file.
class CreatePdfHyperlink
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set the display text for the hyperlink
        sheet.Cells["A1"].PutValue("Open PDF");

        // Add a hyperlink to an external PDF file (opens in a browser)
        // Parameters: cell name, total rows, total columns, address (URL to PDF)
        sheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/document.pdf");

        // Configure HTML save options to open links in a new tab/window
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // _blank target

        // Save the workbook as an HTML file with the specified options
        workbook.Save("PdfLink.html", saveOptions);
    }
}
