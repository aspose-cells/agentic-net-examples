// Title: Convert Excel to HTML with _blank hyperlinks using Aspose.Cells for .NET (C#)
// Description: Shows how to create or load a Workbook, add a hyperlink, configure HtmlSaveOptions with default settings, set LinkTargetType to Blank so links open in a new tab, and save the workbook as an HTML file.
// Keywords: Aspose.Cells | C# | Excel to HTML | HtmlSaveOptions | LinkTargetType | hyperlink target blank | export workbook to HTML | convert Excel to HTML .NET | Aspose.Cells HTML export | open links in new tab
// Common Searches: Aspose.Cells export Excel to HTML C# | HtmlSaveOptions LinkTargetType Blank example | how to make hyperlinks open in new tab when saving Excel as HTML | C# convert workbook to HTML with Aspose.Cells | set hyperlink target _blank Aspose.Cells HTML output
// Developer Intent: Generate an HTML file from an Excel workbook where all hyperlinks open in a new browser tab.
// Use Cases: Create web‑ready reports from Excel data that keep users on the host page. | Automate conversion of dynamically generated spreadsheets for embedding in portals or intranets. | Produce static documentation with external links that open in separate tabs to avoid navigation loss.
// AI Prompts: Provide a C# snippet that converts an Aspose.Cells Workbook to HTML and forces hyperlinks to open in a new tab. | How do I set HtmlSaveOptions.LinkTargetType to Blank for all links during Excel‑to‑HTML conversion with Aspose.Cells? | Explain the steps to export an Excel file to HTML using Aspose.Cells while ensuring hyperlinks use the _blank target.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to create or load a Workbook, add a hyperlink, configure HtmlSaveOptions with default settings, set LinkTargetType to Blank so links open in a new tab, and save the workbook as an HTML file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and a hyperlink
            sheet.Cells["A1"].PutValue("Visit Aspose");
            sheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

            // Create HTML save options with default settings
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set hyperlink target to open in a new window/tab (_blank)
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Save the workbook as HTML
            workbook.Save("output.html", saveOptions);
        }
    }
}
