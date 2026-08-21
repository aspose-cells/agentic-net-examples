// Title: Set All Hyperlink Targets to _blank in Aspose.Cells HTML Export (C#)
// Description: Demonstrates how to configure HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank so that every hyperlink in the generated HTML file opens in a new browser tab.
// Keywords: Aspose.Cells | HtmlSaveOptions | LinkTargetType | HtmlLinkTargetType.Blank | .NET | C# | export Excel to HTML | hyperlink target blank | open links in new tab
// Common Searches: Aspose.Cells set hyperlink target blank | HtmlSaveOptions LinkTargetType example | export Excel as HTML with _blank links | C# Aspose.Cells open links in new tab | how to make HTML links open in new window using Aspose
// Developer Intent: Configure Aspose.Cells to generate HTML where all hyperlinks use target="_blank".
// Use Cases: Publish a financial dashboard online while keeping external references from navigating away. | Create web‑ready documentation from an Excel template with links that open in separate tabs. | Build an interactive spreadsheet portal where users can follow links without losing the main page.
// AI Prompts: Show C# code that sets HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank for an Aspose.Cells workbook. | How can I export an Excel file to HTML with all hyperlinks opening in a new tab using Aspose.Cells? | Explain the effect of HtmlLinkTargetType on generated HTML and how to change it to _blank.

using System;
using Aspose.Cells;

// Demonstrates how to configure HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank so that every hyperlink in the generated HTML file opens in a new browser tab.
class SetHtmlLinkTargetBlank
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put display text into a cell and add a hyperlink to it
        worksheet.Cells["A1"].PutValue("Visit Aspose");
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Create HTML save options and set the link target type to "_blank"
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
