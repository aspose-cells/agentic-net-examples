// Title: Export Excel to HTML without Hidden Rows/Columns – TableCssId on Visible Table (Aspose.Cells C#)
// Description: Demonstrates how to hide specific rows and columns in an Aspose.Cells workbook, configure HtmlSaveOptions to remove those hidden elements, and assign a custom TableCssId that appears only on the generated visible HTML table.
// Keywords: Aspose.Cells export hidden rows HTML | HtmlSaveOptions TableCssId C# | remove hidden columns Aspose.Cells | HtmlHiddenRowDisplayType.Remove example | HtmlHiddenColDisplayType.Remove Aspose | C# Excel to HTML without hidden data | Aspose.Cells HTML table CSS ID
// Common Searches: Aspose.Cells hide row and column then export to HTML | TableCssId property usage in Aspose.Cells HtmlSaveOptions | How to exclude hidden rows/columns when saving Excel as HTML | C# code to export visible cells only with custom CSS ID | Aspose.Cells HTML export remove hidden elements
// Developer Intent: Generate an HTML file from an Excel workbook that contains only visible rows and columns, applying a custom TableCssId solely to the visible table.
// Use Cases: Create web‑ready reports that omit hidden spreadsheet data. | Produce HTML email templates where styling targets only displayed cells. | Publish documentation from Excel while preserving layout without exposing hidden rows or columns.
// AI Prompts: Write C# code using Aspose.Cells to export an Excel worksheet to HTML, removing hidden rows and columns and setting TableCssId to a custom identifier. | Explain the effect of HtmlHiddenRowDisplayType.Remove and HtmlHiddenColDisplayType.Remove on the generated HTML and how TableCssId is applied. | Provide step‑by‑step instructions for hiding rows/columns in a worksheet and exporting only the visible portion to HTML with a custom CSS ID using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to hide specific rows and columns in an Aspose.Cells workbook, configure HtmlSaveOptions to remove those hidden elements, and assign a custom TableCssId that appears only on the generated visible HTML table.
class ExportHiddenRowsAndColumns
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");
        worksheet.Cells["C2"].PutValue("Data3");

        // Hide row 2 (index 1) and column B (index 1)
        worksheet.Cells.HideRow(1);
        worksheet.Cells.HideColumn(1);

        // Set HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Prefix for CSS classes inside the generated table element
            TableCssId = "visibleTable",
            // Remove hidden rows and columns from the HTML output
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove,
            HiddenColDisplayType = HtmlHiddenColDisplayType.Remove
        };

        // Save the workbook as HTML; only visible rows/columns will be exported
        workbook.Save("output.html", saveOptions);
    }
}
