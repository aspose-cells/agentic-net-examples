// Title: Export Excel to a Single HTML File with JavaScript Page‑Number Footer using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills sample cells, defines a centered footer with &P (current page) and &N (total pages) placeholders, and saves the workbook as one HTML file using HtmlSaveOptions that enable page footers, headers and JavaScript compatibility.
// Keywords: Aspose.Cells | C# | .NET | Export Excel to HTML | HTML footer page numbers | SaveAsSingleFile | ExportPageFooters | JavaScript pagination | Excel web report | PageSetup footer
// Common Searches: Aspose.Cells export Excel to single HTML file with footer | Add page numbers to HTML export using Aspose.Cells .NET | Enable JavaScript‑compatible HTML output with page footers Aspose | Set footer text &P &N in Aspose.Cells HTML export | C# export workbook to HTML with page numbering
// Developer Intent: Generate a single‑file HTML export of an Excel workbook that displays the current page and total page count in the footer via JavaScript.
// Use Cases: Web‑based printable reports that need "Page X of Y" footers | Embedding Excel data in a web app while preserving pagination information | Creating HTML invoices or statements that require legal page numbering
// AI Prompts: Show how to align the footer text to the right while keeping page numbers in the exported HTML. | Provide code to attach a custom CSS stylesheet to the HTML output without breaking the JavaScript footer. | Explain how to export each worksheet to separate HTML files, each with a page‑number footer, using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, fills sample cells, defines a centered footer with &P (current page) and &N (total pages) placeholders, and saves the workbook as one HTML file using HtmlSaveOptions that enable page footers, headers and JavaScript compatibility.
class ExportExcelToHtmlWithFooter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Jane");
        worksheet.Cells["B3"].PutValue(28);

        // Set the footer: center section will display page number and total pages
        // &P = current page number, &N = total page count
        worksheet.PageSetup.SetFooter(1, "Page &P of &N");

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportPageFooters = true;   // enable footer export
        saveOptions.ExportPageHeaders = true;   // optional: export headers as well
        saveOptions.SaveAsSingleFile = true;    // required for footer export
        saveOptions.IsJsBrowserCompatible = true; // ensure JavaScript works in browsers

        // Save the workbook as an HTML file with the configured options
        workbook.Save("WorkbookWithFooter.html", saveOptions);
    }
}
