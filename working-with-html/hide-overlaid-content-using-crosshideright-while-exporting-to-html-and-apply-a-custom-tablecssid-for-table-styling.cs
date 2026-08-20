// Title: Hide Overlapping Cell Text with CrossHideRight & Apply Custom TableCssId in Aspose.Cells HTML Export (C#)
// Description: Shows how to create a workbook, insert a long string that overflows into the next cell, and configure HtmlSaveOptions to use HtmlCrossType.CrossHideRight to hide the right part of the overflow while assigning TableCssId="custom-table-style" for external CSS styling, then save the sheet as HTML.
// Keywords: Aspose.Cells | HTML export | CrossHideRight | HtmlCrossType | TableCssId | C# | .NET | hide overlapping text | custom table CSS | workbook to HTML | export options
// Common Searches: Aspose.Cells hide overlapping text HTML | CrossHideRight example C# | Set TableCssId in HtmlSaveOptions | C# export workbook to HTML with custom CSS id | truncate overflow cell text Aspose.Cells
// Developer Intent: Export a workbook to HTML while truncating the right‑hand overflow of long cell values and assigning a specific CSS id to the generated table for styling.
// Use Cases: Generate HTML reports where long cell values must not spill into adjacent columns. | Apply external stylesheet rules to the exported table via a known CSS id. | Create printable HTML output with consistent layout and controlled overflow handling. | Integrate HTML export into web applications that require precise table styling.
// AI Prompts: Provide C# code using Aspose.Cells to export a worksheet to HTML with HtmlCrossType.CrossHideRight and TableCssId='custom-table-style'. | Explain how HtmlCrossType.CrossHideRight modifies overlapping cell content in the generated HTML and how to style the table using the assigned TableCssId. | Recommend additional HtmlSaveOptions (e.g., EmbedCss, ExportImagesAsBase64, PreserveFormatting) to improve HTML export quality while keeping the custom table id.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Shows how to create a workbook, insert a long string that overflows into the next cell, and configure HtmlSaveOptions to use HtmlCrossType.CrossHideRight to hide the right part of the overflow while assigning TableCssId="custom-table-style" for external CSS styling, then save the sheet as HTML.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells with data that will cause a cross‑cell string
            // Cell A1 contains a long text that will overflow into B1
            sheet.Cells["A1"].PutValue("This is a very long text that will overlap the next cell when rendered.");
            sheet.Cells["B1"].PutValue("Right side");

            // Optional: set column widths to make the overlap visible
            sheet.Cells.SetColumnWidth(0, 20); // Column A
            sheet.Cells.SetColumnWidth(1, 15); // Column B

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Hide the right part of overlapping text
            saveOptions.HtmlCrossStringType = HtmlCrossType.CrossHideRight;

            // Apply a custom TableCssId for styling the generated HTML table
            saveOptions.TableCssId = "custom-table-style";

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedWithCrossHideRight.html", saveOptions);
        }
    }
}
