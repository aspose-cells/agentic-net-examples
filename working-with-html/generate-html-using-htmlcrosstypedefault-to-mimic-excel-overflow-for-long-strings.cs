// Title: Export Excel to HTML with Excel‑style overflow via HtmlCrossType.Default (C# Aspose.Cells)
// Description: Demonstrates how to create a workbook, insert a long string in cell A1, set a narrow column width, configure HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.Default, and save the sheet as HTML so the text overflows like in Excel.
// Keywords: Aspose.Cells HTML export | HtmlCrossType.Default | Excel overflow HTML | C# Aspose.Cells example | cell text overflow | HtmlSaveOptions cross string | export worksheet to HTML
// Common Searches: Aspose.Cells export HTML overflow behavior | HtmlCrossStringType Default example C# | how to keep Excel text overflow in HTML output | C# save workbook as HTML with overflow | Aspose.Cells HTMLCrossType settings
// Developer Intent: Export a spreadsheet to HTML while preserving Excel‑like text overflow using HtmlCrossType.Default.
// Use Cases: Generate web‑ready reports that display long cell values spilling into adjacent columns. | Create an online spreadsheet viewer that mimics Excel’s visual layout for better readability. | Produce printable HTML previews where column width constraints and overflow are retained.
// AI Prompts: Show C# code to export multiple worksheets to HTML with HtmlCrossType.Default so each sheet keeps overflow behavior. | Explain how column width units influence overflow when using HtmlCrossType.Default in Aspose.Cells. | Provide a comparison of HtmlCrossType.Default vs. HtmlCrossType.Always for controlling text overflow in HTML exports.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossDemo
{
    // Demonstrates how to create a workbook, insert a long string in cell A1, set a narrow column width, configure HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.Default, and save the sheet as HTML so the text overflows like in Excel.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into a cell that will overflow the column width
            sheet.Cells["A1"].PutValue("This is a very long text string that should overflow the cell width and demonstrate Excel-like overflow behavior when exported to HTML.");

            // Set a narrow column width to force overflow
            sheet.Cells.SetColumnWidth(0, 10); // width in characters

            // Configure HTML save options to use the Default cross‑string behavior
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Default; // mimic Excel overflow

            // Save the workbook as an HTML file
            workbook.Save("OverflowDemo.html", htmlOptions);
        }
    }
}
