// Title: Aspose.Cells C# – Export Excel to HTML with hyperlinks opening in a new tab
// Description: Shows how to create a workbook, add a hyperlink, and set HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank so the saved HTML file uses target="_blank" for every link.
// Keywords: Aspose.Cells | HtmlSaveOptions | HtmlLinkTargetType | Blank | C# | open link in new tab | export Excel to HTML | hyperlink target | Aspose.Cells example | HTML export options
// Common Searches: Aspose.Cells set hyperlink target _blank | HtmlSaveOptions LinkTargetType C# | Export Excel as HTML with links opening in new tab | Aspose.Cells HTML export hyperlink target | How to make Aspose.Cells HTML links open in new window
// Developer Intent: Configure HTML export so all workbook hyperlinks open in a new browser tab.
// Use Cases: Creating web‑ready reports from Excel where external links should not replace the current page. | Generating HTML documentation from spreadsheets with consistent link behavior. | Building email templates from Excel that require links to open in a separate tab.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML with LinkTargetType set to Blank. | Explain the effect of HtmlLinkTargetType.Blank on the generated HTML file. | Show an example of adding a hyperlink to a cell and exporting to HTML so the link opens in a new tab.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkTargetDemo
{
    // Shows how to create a workbook, add a hyperlink, and set HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank so the saved HTML file uses target="_blank" for every link.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some display text into a cell
            sheet.Cells["A1"].PutValue("Visit Aspose");

            // Add a hyperlink to the cell (A1)
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.aspose.com");

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set the link target type to open in a new tab/window (_blank)
            saveOptions.LinkTargetType = HtmlLinkTargetType.Blank;

            // Save the workbook as an HTML file with the configured options
            workbook.Save("HyperlinkTargetBlank.html", saveOptions);
        }
    }
}
