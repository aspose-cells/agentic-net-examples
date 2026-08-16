// Title: Insert a hyperlink that opens in a new browser tab using Aspose.Cells for .NET (C#)
// Description: Shows how to add a hyperlink to a worksheet cell, set its display text and screen tip, configure HtmlSaveOptions.LinkTargetType to Blank (target="_blank"), and save the workbook as HTML so the link opens in a new tab.
// Keywords: Aspose.Cells C# hyperlink | HtmlSaveOptions target blank | add hyperlink to Excel cell | open link in new tab Aspose | save workbook as HTML | hyperlink screen tip | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add hyperlink target _blank | C# Aspose.Cells hyperlink open new tab | How to set HtmlSaveOptions.LinkTargetType in Aspose.Cells | Create hyperlink with screen tip in Aspose.Cells | Export Excel to HTML with clickable links Aspose
// Developer Intent: Add a cell hyperlink that points to an external URL and ensure it opens in a new browser tab when the workbook is exported to HTML.
// Use Cases: Generate an HTML report where each link opens external documentation in a separate tab. | Build a marketing dashboard with product links that launch in new windows for uninterrupted browsing. | Provide additional context via screen tips on hyperlinks in exported HTML files.
// AI Prompts: Write C# code with Aspose.Cells to insert a hyperlink in cell B2, display text "Documentation", and open it in a new tab when saved as HTML. | Explain how HtmlSaveOptions.LinkTargetType = Blank affects hyperlink behavior in the generated HTML file. | Show how to add a screen tip to a hyperlink in Aspose.Cells and export the workbook to HTML with target="_blank".

using System;
using Aspose.Cells;

// Shows how to add a hyperlink to a worksheet cell, set its display text and screen tip, configure HtmlSaveOptions.LinkTargetType to Blank (target="_blank"), and save the workbook as HTML so the link opens in a new tab.
class InsertHyperlinkExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A1 that points to an external website
        int linkIndex = sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

        // Set the display text for the hyperlink
        sheet.Hyperlinks[linkIndex].TextToDisplay = "Visit Example";

        // Optional: add a screen tip
        sheet.Hyperlinks[linkIndex].ScreenTip = "Open Example website";

        // Configure HTML save options so that links open in a new tab/window
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank; // corresponds to target="_blank"

        // Save the workbook as HTML; the hyperlink will open in a new tab when clicked
        workbook.Save("HyperlinkExample.html", htmlOptions);
    }
}
