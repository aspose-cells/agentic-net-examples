// Title: Aspose.Cells C# – Set all HTML hyperlink targets to "_blank" using HtmlSaveOptions.LinkTargetType
// Description: Demonstrates how to create a workbook, add a hyperlink, configure HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank, and save the file as HTML so every link opens in a new browser tab.
// Keywords: Aspose.Cells HtmlSaveOptions | HtmlLinkTargetType.Blank | C# export Excel to HTML | hyperlink target _blank | Aspose.Cells hyperlink settings
// Common Searches: Aspose.Cells set hyperlink target blank | HtmlSaveOptions LinkTargetType example C# | export Excel to HTML with _blank links | make all links open in new tab Aspose.Cells
// Developer Intent: Configure Aspose.Cells HTML export so that all hyperlinks are rendered with target="_blank".
// Use Cases: Publish a spreadsheet on a website where external links must not navigate away from the page. | Generate an HTML report for internal users, ensuring links open in separate tabs. | Create an HTML email attachment from a workbook with links that open in new windows.
// AI Prompts: Show how to set HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank for every hyperlink when saving a workbook to HTML with Aspose.Cells (C#). | Provide a complete C# sample that adds multiple hyperlinks to a workbook and exports it to HTML with target='_blank' on all links. | Explain the differences between HtmlLinkTargetType.Blank, Self, and Parent in Aspose.Cells and recommend scenarios for each.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add a hyperlink, configure HtmlSaveOptions.LinkTargetType to HtmlLinkTargetType.Blank, and save the file as HTML so every link opens in a new browser tab.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put display text into cell A1
        sheet.Cells["A1"].PutValue("Visit Google");

        // Add a hyperlink to cell A1 (display text -> URL)
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.google.com");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Set the target type so that all links are rendered with target="_blank"
        saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
