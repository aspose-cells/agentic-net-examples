using System;
using Aspose.Cells;

class RtlToHtmlExport
{
    static void Main()
    {
        // Load the workbook that contains right‑to‑left (RTL) text
        Workbook workbook = new Workbook("input.xlsx");

        // Enable RTL display for all worksheets
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.DisplayRightToLeft = true;
        }

        // Configure HTML export options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Show the full cell content even if it exceeds column width
        htmlOptions.FormatDataIgnoreColumnWidth = true;
        // Use HTML5 for better compatibility (optional)
        htmlOptions.HtmlVersion = HtmlVersion.Html5;

        // Export the workbook to HTML
        workbook.Save("output.html", htmlOptions);
    }
}