// Title: Insert a PDF hyperlink in Aspose.Cells (C#) and export to HTML with a _blank target
// Description: Demonstrates how to add a hyperlink to an external PDF in cell A1, set its display text and screen tip, configure HtmlSaveOptions.LinkTargetType to open links in a new browser tab, and save the workbook as an HTML file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF hyperlink | C# HtmlSaveOptions LinkTargetType Blank | open hyperlink in new tab | set hyperlink TextToDisplay | set hyperlink ScreenTip | export Excel to HTML Aspose | .NET Aspose.Cells hyperlink example
// Common Searches: Aspose.Cells add hyperlink to PDF and open in new tab | C# HtmlSaveOptions LinkTargetType Blank example | How to set screen tip for hyperlink in Aspose.Cells | Export Excel workbook to HTML with _blank links | Aspose.Cells create clickable PDF links in worksheet
// Developer Intent: Add a PDF hyperlink to a worksheet, customize its label and tooltip, and generate HTML where the link opens in a new browser tab.
// Use Cases: Generate web‑ready reports that reference PDF manuals without leaving the page. | Create an online product catalog where each item links to a downloadable brochure. | Publish Excel‑based dashboards with external documentation links that open in separate tabs.
// AI Prompts: Write C# code with Aspose.Cells to insert a PDF hyperlink in cell A1, set TextToDisplay and ScreenTip, configure HtmlSaveOptions to use the _blank target, and save as HTML. | Explain the effect of HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank on hyperlink behavior in exported HTML. | Show how to retrieve a newly added hyperlink from a worksheet and modify its properties using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExample
{
    // Demonstrates how to add a hyperlink to an external PDF in cell A1, set its display text and screen tip, configure HtmlSaveOptions.LinkTargetType to open links in a new browser tab, and save the workbook as an HTML file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a hyperlink to cell A1 that points to an external PDF file
            // Parameters: cell name, total rows, total columns, address (URL of the PDF)
            int hyperlinkIndex = worksheet.Hyperlinks.Add("A1", 1, 1, "https://example.com/document.pdf");

            // Retrieve the created hyperlink to set display text and screen tip
            Hyperlink hyperlink = worksheet.Hyperlinks[hyperlinkIndex];
            hyperlink.TextToDisplay = "Open PDF Document";
            hyperlink.ScreenTip = "Click to view the PDF in a new tab";

            // Configure HTML save options to open links in a new browser tab/window
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank; // _blank target

            // Save the workbook as an HTML file using the configured options
            workbook.Save("HyperlinkToPdf.html", saveOptions);
        }
    }
}
