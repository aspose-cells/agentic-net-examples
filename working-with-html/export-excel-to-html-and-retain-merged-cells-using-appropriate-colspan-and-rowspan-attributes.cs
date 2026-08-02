// Title: Export Excel to HTML with merged cells using Aspose.Cells for .NET
// Description: Shows how to build a workbook, merge a range, and save it as HTML with Aspose.Cells. The HtmlSaveOptions generate the correct colspan/rowspan tags so the merged layout is retained in the web page.
// Keywords: Aspose.Cells | C# | .NET | HTML export | merged cells | colspan | rowspan | HtmlSaveOptions | MergeEmptyTdType | Excel to HTML conversion
// Common Searches: Aspose.Cells export merged cells to HTML | HTML output with colspan from Excel using Aspose | keep merged cells when converting Excel to HTML .NET | HtmlSaveOptions MergeEmptyTdType default behavior | C# code to convert Excel with merged headers to HTML
// Developer Intent: Convert an Excel workbook to an HTML document while preserving merged ranges through proper table attributes.
// Use Cases: Render Excel reports with merged header rows as web‑ready HTML tables. | Automate the transformation of dashboard worksheets into static HTML pages without losing layout. | Create documentation that mirrors the original spreadsheet’s merged cell structure in a browser.
// AI Prompts: Provide C# code that exports an Aspose.Cells workbook to HTML and keeps merged cells using colspan and rowspan. | Explain how HtmlSaveOptions.MergeEmptyTdType influences the rendering of empty merged cells in HTML. | Show how to customize the HTML output of merged cells with additional CSS styling in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to build a workbook, merge a range, and save it as HTML with Aspose.Cells. The HtmlSaveOptions generate the correct colspan/rowspan tags so the merged layout is retained in the web page.
class ExportMergedCellsToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data
        sheet.Cells["A1"].PutValue("Header 1");
        sheet.Cells["B1"].PutValue("Header 2");
        sheet.Cells["A2"].PutValue("Value 1");
        sheet.Cells["B2"].PutValue("Value 2");

        // Merge cells A1:B1 to demonstrate merged cells (will be rendered with colspan)
        sheet.Cells.Merge(0, 0, 1, 2); // Row 0, Column 0, 1 row, 2 columns

        // Set HTML save options – default merging behavior retains merged cells
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            MergeEmptyTdType = MergeEmptyTdType.Default
        };

        // Save the workbook as HTML; merged cells will appear with appropriate colspan/rowspan
        workbook.Save("MergedCells.html", htmlOptions);
    }
}
