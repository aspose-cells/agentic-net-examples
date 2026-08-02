// Title: C# AspNet – Export Excel to HTML while preserving wrapped text with Aspose.Cells
// Description: Shows how to build a workbook, turn on text wrapping for a cell, narrow the column, set HtmlSaveOptions.HideOverflowWrappedText = false, and save the file as HTML so the wrapped content remains visible.
// Keywords: Aspose.Cells | C# | .NET | HTML export | wrap text | HideOverflowWrappedText | cell style | StyleFlag | Excel to HTML | preserve wrapped text | generate HTML report
// Common Searches: Aspose.Cells export HTML keep wrapped text visible | C# save Excel as HTML without clipping wrapped cells | HtmlSaveOptions HideOverflowWrappedText example | how to enable text wrap in Aspose.Cells HTML output | C# generate HTML from Excel with column width and wrap
// Developer Intent: Create an HTML file from an Excel workbook that retains cell text‑wrapping formatting.
// Use Cases: Building web‑based reports where long descriptions must wrap inside narrow columns. | Embedding Excel‑styled tables in newsletters or intranet pages without losing wrap formatting. | Providing a printable HTML view of a spreadsheet where overflow text should stay on screen.
// AI Prompts: Give a code snippet that adds a custom CSS class to cells with wrapped text in the generated HTML. | Show how to export several worksheets to one HTML document while keeping wrap settings for all cells. | Explain how to auto‑adjust column widths based on the length of wrapped content before saving to HTML.

using System;
using Aspose.Cells;

// Shows how to build a workbook, turn on text wrapping for a cell, narrow the column, set HtmlSaveOptions.HideOverflowWrappedText = false, and save the file as HTML so the wrapped content remains visible.
class GenerateHtmlWithWrap
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Insert a long text into cell A1
        cells["A1"].PutValue("This is a long text that should wrap inside the cell when exported to HTML. It contains enough characters to demonstrate wrapping behavior.");

        // Enable text wrapping for the cell using a style and a style flag
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;
        StyleFlag wrapFlag = new StyleFlag();
        wrapFlag.WrapText = true;
        cells["A1"].SetStyle(wrapStyle, wrapFlag);

        // Set column width to a narrow value so the text needs to wrap
        cells.SetColumnWidth(0, 15); // width in characters

        // Configure HTML save options to preserve wrapped text (do not hide overflow)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.HideOverflowWrappedText = false; // ensure overflow text remains visible

        // Save the workbook as an HTML file with the specified options
        workbook.Save("WrappedTextOutput.html", htmlOptions);
    }
}
