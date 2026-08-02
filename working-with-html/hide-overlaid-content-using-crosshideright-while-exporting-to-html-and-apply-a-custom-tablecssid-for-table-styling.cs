// Title: Hide Overlaid Text with CrossHideRight and Apply a Custom TableCssId When Exporting Excel to HTML – Aspose.Cells for .NET
// Description: Shows how to export an Aspose.Cells workbook to HTML while using HtmlCrossType.CrossHideRight to suppress right‑hand overflow text and assigning a TableCssId to style the generated table with custom CSS.
// Keywords: Aspose.Cells | HtmlSaveOptions | CrossHideRight | HtmlCrossType | TableCssId | export Excel to HTML | .NET | overlay text handling | custom CSS id | HTML table styling
// Common Searches: Aspose.Cells hide overflow text in HTML export | CrossHideRight example C# | Set TableCssId for HTML table Aspose.Cells | Export Excel workbook to HTML with custom CSS id | Prevent cell text overlapping in HTML output Aspose
// Developer Intent: Export a workbook to HTML, hide the right portion of overlaid cell text, and give the output table a specific CSS identifier for styling.
// Use Cases: Creating HTML reports where long cell values must not spill into adjacent cells. | Applying brand‑consistent CSS rules to the exported table by using a custom TableCssId. | Generating printable HTML versions of spreadsheets with controlled overlay behavior for better layout control.
// AI Prompts: How do I configure HtmlSaveOptions to use CrossHideRight and set a custom TableCssId in Aspose.Cells for .NET? | Provide a C# snippet that exports an Excel workbook to HTML with hidden right‑hand overflow text and a custom table CSS id. | Explain the effect of HtmlCrossType.CrossHideRight on overlaid text in the HTML output and how to verify it.

using System;
using Aspose.Cells;

// Shows how to export an Aspose.Cells workbook to HTML while using HtmlCrossType.CrossHideRight to suppress right‑hand overflow text and assigning a TableCssId to style the generated table with custom CSS.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add data that will cause overlaid (cross‑cell) text
        // Long text in A1 will overflow into B1 when column width is small
        sheet.Cells["A1"].PutValue("This is a very long text that will overflow into the next cell");
        sheet.Cells["B1"].PutValue("Right side content");

        // Reduce column widths to force overflow
        sheet.Cells.SetColumnWidth(0, 10); // Column A
        sheet.Cells.SetColumnWidth(1, 10); // Column B

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Hide the right part of overlaid text when it overlaps
        htmlOptions.HtmlCrossStringType = HtmlCrossType.CrossHideRight;
        // Apply a custom TableCssId for styling the generated table
        htmlOptions.TableCssId = "custom-table-style";

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
