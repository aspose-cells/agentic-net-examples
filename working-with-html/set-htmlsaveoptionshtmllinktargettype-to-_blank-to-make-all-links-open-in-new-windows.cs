// Title: Aspose.Cells C# – Open all HTML hyperlinks in a new tab with HtmlSaveOptions.LinkTargetType=_blank
// Description: Learn how to configure Aspose.Cells HtmlSaveOptions so that every hyperlink in the exported HTML file uses target="_blank". The example creates a workbook, adds a hyperlink, sets LinkTargetType to HtmlLinkTargetType.Blank, and saves the file, ensuring links open in a new browser window or tab.
// Keywords: Aspose.Cells HtmlSaveOptions | LinkTargetType Blank | C# export Excel to HTML | target=_blank hyperlinks | Aspose.Cells hyperlink new tab | HTML export options Aspose | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set link target to _blank | C# HtmlSaveOptions LinkTargetType example | export Excel to HTML with new tab links | Aspose.Cells hyperlink target attribute | how to open Aspose.Cells HTML links in new window
// Developer Intent: Configure Aspose.Cells HTML export so that all hyperlinks open in a new browser tab or window.
// Use Cases: Publishing financial dashboards where external references should not replace the main report view. | Generating web‑ready documentation from Excel with links that keep the documentation page open. | Embedding spreadsheet data in a portal while ensuring safe navigation to external resources.
// AI Prompts: Show C# code to set HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank in Aspose.Cells. | Provide an Aspose.Cells example that adds a hyperlink and saves the workbook as HTML with target=_blank. | Explain HtmlLinkTargetType options (Blank, Self, Parent, Top) in Aspose.Cells HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsLinkTargetDemo
{
    // Learn how to configure Aspose.Cells HtmlSaveOptions so that every hyperlink in the exported HTML file uses target="_blank". The example creates a workbook, adds a hyperlink, sets LinkTargetType to HtmlLinkTargetType.Blank, and saves the file, ensuring links open in a new browser window or tab.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample text and a hyperlink in cell A1
            sheet.Cells["A1"].PutValue("Visit Aspose");
            sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set the link target type to "_blank" so links open in a new window/tab
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Save the workbook as HTML using the configured options
            workbook.Save("AsposeLinkTargetBlank.html", saveOptions);
        }
    }
}
