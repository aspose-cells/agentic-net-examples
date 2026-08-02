// Title: Export Excel to HTML with clickable hyperlinks using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a cell value and a hyperlink, configure HtmlSaveOptions to set LinkTargetType to Blank, and save the file as HTML so the link opens in a new browser tab and remains functional.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | hyperlink export | HtmlSaveOptions | LinkTargetType | clickable links | browser-friendly HTML | preserve Excel hyperlinks
// Common Searches: Aspose.Cells export Excel to HTML with links | C# save workbook as HTML clickable hyperlinks | HtmlSaveOptions LinkTargetType blank example | how to keep Excel hyperlinks in HTML output | Aspose.Cells HTML export hyperlink target
// Developer Intent: Generate an HTML file from an Excel workbook where all cell hyperlinks stay active and open in a new tab.
// Use Cases: Publish a product catalog with direct links to item pages as a web‑ready HTML file. | Distribute a data report that includes references to external documentation via clickable URLs. | Create an email‑compatible HTML version of a spreadsheet that retains functional hyperlinks.
// AI Prompts: Write C# code with Aspose.Cells to add several hyperlinks to different cells and export the workbook to HTML, ensuring each link opens in a new window. | Explain the effect of HtmlSaveOptions.LinkTargetType on generated HTML and show how to switch between '_self' and '_blank' targets. | Provide a step‑by‑step tutorial for converting an Excel file to HTML while preserving all hyperlinks and customizing their target attributes.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, insert a cell value and a hyperlink, configure HtmlSaveOptions to set LinkTargetType to Blank, and save the file as HTML so the link opens in a new browser tab and remains functional.
class ExportExcelToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to a cell
        worksheet.Cells["A1"].PutValue("Visit Google");

        // Add a hyperlink to the same cell (display text will be the cell value)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
        worksheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.google.com");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Ensure hyperlinks are functional in the browser.
        // Setting the target type to "_blank" opens links in a new tab/window.
        htmlOptions.LinkTargetType = HtmlLinkTargetType.Blank;

        // Save the workbook as an HTML file with the specified options
        workbook.Save("ExportedWorkbook.html", htmlOptions);
    }
}
