// Title: Export Worksheet to Single HTML with Styles & Merged Cells using Aspose.Cells (C#)
// Description: Demonstrates how to save the active worksheet of an Aspose.Cells workbook as a single HTML file while preserving cell formatting, merged ranges, grid lines, and embedding CSS. The example uses HtmlSaveOptions to control output.
// Keywords: Aspose.Cells HTML export | C# export worksheet to HTML | preserve merged cells Aspose | cell style HTML Aspose.Cells | HtmlSaveOptions single file | embed CSS Aspose.Cells | grid lines HTML export | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells export worksheet to HTML with merged cells | How to keep cell formatting when saving Excel as HTML in C# | Save Aspose.Cells workbook as one HTML file | Embed CSS in HTML output from Aspose.Cells | Export active sheet only Aspose.Cells HTML
// Developer Intent: Create an HTML snapshot of the active worksheet that looks identical to the Excel view, including styles and merged cells, in a single file.
// Use Cases: Generate printable web reports that match the original spreadsheet layout. | Provide a quick, styled preview of a dashboard worksheet on a website. | Attach a fully formatted worksheet snapshot to an email without sending the Excel file.
// AI Prompts: Show how to modify HtmlSaveOptions to output CSS to an external .css file while preserving merged cells. | Provide code that exports every worksheet in a workbook to separate HTML files, keeping all formatting. | Explain how to include embedded images from the workbook when exporting to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// Demonstrates how to save the active worksheet of an Aspose.Cells workbook as a single HTML file while preserving cell formatting, merged ranges, grid lines, and embedding CSS. The example uses HtmlSaveOptions to control output.
class ExportWorksheetToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data and apply a style to a cell
        sheet.Cells["A1"].PutValue("Header");
        Style headerStyle = sheet.Cells["A1"].GetStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightBlue;
        headerStyle.Pattern = BackgroundType.Solid;
        sheet.Cells["A1"].SetStyle(headerStyle);

        // Merge cells A1:C1 to preserve merged cell structure in HTML
        sheet.Cells.Merge(0, 0, 1, 3); // row 0, column 0, 1 row, 3 columns

        // Add additional data
        sheet.Cells["A2"].PutValue("Item");
        sheet.Cells["B2"].PutValue(123);
        sheet.Cells["C2"].PutValue(DateTime.Now);

        // Configure HTML save options to keep styles and merged cells
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.ExportActiveWorksheetOnly = true;          // Export only the active sheet
        options.ExportWorksheetProperties = true;         // Preserve worksheet properties
        options.ExportWorksheetCSSSeparately = false;     // Embed CSS in the HTML file
        options.SaveAsSingleFile = true;                  // Produce a single HTML file
        options.ExportGridLines = true;                   // Optional: show grid lines

        // Define output path (e.g., Desktop)
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "WorksheetExport.html");

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, options);

        Console.WriteLine("Worksheet exported to HTML: " + outputPath);
    }
}
