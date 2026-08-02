// Title: C# Aspose.Cells – Save Workbook as HTML with HtmlCrossType.Default to mimic Excel overflow
// Description: Shows how to create a workbook, put a long string in a cell, narrow the column, set HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.Default, and save as HTML so the text flows into neighboring cells just like Excel.
// Keywords: Aspose.Cells | HtmlCrossType.Default | HTML export | Excel overflow | C# example | long text overflow | HtmlSaveOptions | cross type | cell overflow | Excel-like HTML
// Common Searches: Aspose.Cells HtmlCrossType.Default overflow example | export Excel to HTML preserving text overflow | C# save workbook as HTML with Excel-like overflow | how to keep long cell values visible in HTML using Aspose.Cells | HtmlSaveOptions HtmlCrossStringType usage
// Developer Intent: Export a worksheet to HTML while keeping the Excel‑style overflow of long cell values using HtmlCrossType.Default.
// Use Cases: Generate HTML reports that display long strings across adjacent empty cells, matching the Excel view. | Create web‑based dashboards where column widths are fixed but text should flow without truncation. | Provide printable HTML previews of spreadsheets that retain the original overflow appearance.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as HTML using HtmlCrossType.Default so that long text overflows adjacent cells. | Explain how HtmlCrossStringType influences the HTML output and how to configure it for Excel‑like overflow. | Give a step‑by‑step guide to set a narrow column width, insert a long string, and export to HTML with overflow using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossDemo
{
    // Shows how to create a workbook, put a long string in a cell, narrow the column, set HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.Default, and save as HTML so the text flows into neighboring cells just like Excel.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a long text that will overflow the cell width
            sheet.Cells["A1"].PutValue("This is a very long text string that should overflow the cell width and demonstrate Excel-like overflow behavior.");

            // Set a narrow column width to force overflow
            sheet.Cells.SetColumnWidth(0, 5); // Column A width

            // Configure HTML save options to use the Default cross type (Excel-like overflow)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Default;

            // Save the workbook as an HTML file
            workbook.Save("OverflowDemo.html", htmlOptions);

            Console.WriteLine("HTML file saved with HtmlCrossType.Default to mimic Excel overflow.");
        }
    }
}
