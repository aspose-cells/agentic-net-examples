// Title: Export Excel to HTML with clickable hyperlinks using Aspose.Cells for .NET (C#)
// Description: Learn how to convert an Excel workbook to an HTML page with Aspose.Cells while preserving hyperlinks. The example shows adding a link, configuring HtmlSaveOptions to open links in a new tab (LinkTargetType.Blank) and to use relative URLs (IsFullPathLink = false), then saving the result as a fully functional HTML file.
// Keywords: Aspose.Cells export Excel to HTML | C# preserve hyperlinks HTML | HtmlSaveOptions LinkTargetType Blank | .NET convert Excel to HTML | relative hyperlink paths Aspose.Cells | open links in new browser tab
// Common Searches: Aspose.Cells export Excel to HTML with clickable links | How to keep hyperlinks when saving Excel as HTML .NET | HtmlSaveOptions open hyperlink in new tab C# | relative vs absolute links Aspose.Cells HTML export | C# convert spreadsheet to web‑ready HTML
// Developer Intent: Create an HTML version of an Excel file where all embedded hyperlinks stay active and open in a new browser tab.
// Use Cases: Generate web‑ready reports that include external resource links. | Publish Excel‑based documentation as HTML with navigation links. | Provide intranet dashboards where cells link to internal pages.
// AI Prompts: Show me C# code to export an Excel workbook to HTML with functional hyperlinks using Aspose.Cells, opening links in a new tab. | How do I configure HtmlSaveOptions in Aspose.Cells to use relative hyperlink paths and preserve clickability? | Explain the difference between HtmlLinkTargetType.Blank and other target types when saving Excel as HTML.

using System;
using Aspose.Cells;

// Learn how to convert an Excel workbook to an HTML page with Aspose.Cells while preserving hyperlinks. The example shows adding a link, configuring HtmlSaveOptions to open links in a new tab (LinkTargetType.Blank) and to use relative URLs (IsFullPathLink = false), then saving the result as a fully functional HTML file.
class ExportExcelToHtml
{
    static void Main()
    {
        // Load the source Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure there is a hyperlink in the worksheet (optional example)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
        workbook.Worksheets[0].Hyperlinks.Add(0, 0, 1, 1, "https://www.example.com");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Open hyperlinks in a new browser tab/window
        htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Use relative links (default). Set to true if absolute paths are required.
        htmlOptions.IsFullPathLink = false;

        // Save the workbook as an HTML file with functional hyperlinks
        workbook.Save("output.html", htmlOptions);
    }
}
