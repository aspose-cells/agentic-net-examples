// Title: Export Active Worksheet to HTML with a Custom TableCssId – Aspose.Cells for .NET
// Description: Demonstrates how to save only the active worksheet of a workbook as an HTML file and assign a specific TableCssId to style the generated tables using Aspose.Cells for .NET.
// Keywords: Aspose.Cells export active worksheet HTML | HtmlSaveOptions TableCssId | C# export single sheet to HTML | custom CSS ID for Aspose.Cells tables | Aspose.Cells HTML styling options
// Common Searches: export only active sheet to html asp.net | set tablecssid aspocells htmlsaveoptions | aspocells save active worksheet as html example | how to style html tables from aspocells workbook | c# aspocells export single worksheet html
// Developer Intent: Generate an HTML file that contains only the currently active worksheet and apply a user‑defined CSS ID to its tables for external styling.
// Use Cases: Create a lightweight web report that shows just the active sheet from a multi‑sheet workbook. | Apply a unique CSS identifier to tables so they can be styled consistently across a web application. | Provide a quick preview of a specific worksheet without exporting the entire workbook.
// AI Prompts: Write C# code that saves only the active worksheet to HTML and sets TableCssId to "my-table" with Aspose.Cells. | Explain the effect of ExportActiveWorksheetOnly and TableCssId on the HTML output produced by Aspose.Cells. | Show how to link an external stylesheet to tables generated from a workbook saved as HTML using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to save only the active worksheet of a workbook as an HTML file and assign a specific TableCssId to style the generated tables using Aspose.Cells for .NET.
class ExportActiveWorksheetHtml
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].PutValue("Data in active sheet");

        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].PutValue("Data in other sheet");

        // Set the first worksheet as the active sheet
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Configure HTML save options:
        // - Export only the active worksheet
        // - Apply a custom TableCssId for styling tables
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportActiveWorksheetOnly = true;
        saveOptions.TableCssId = "custom-table-style";

        // Save the workbook to HTML using the configured options
        workbook.Save("active_sheet.html", saveOptions);
    }
}
