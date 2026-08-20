// Title: C# – Export Numeric Cells as Plain Text (No Scientific Notation) with Aspose.Cells HtmlSaveOptions
// Description: Demonstrates how to create a workbook, apply the text format "@" to numeric cells, configure HtmlSaveOptions (ExportDataOptions.All) and save to HTML so large integers and tiny decimals are rendered as plain text instead of scientific notation.
// Keywords: Aspose.Cells | HtmlSaveOptions | C# | export numeric as text | prevent scientific notation | text format @ | HtmlExportDataOptions.All | HTML report generation | cell style text | large numbers | small decimals
// Common Searches: Aspose.Cells export numbers as plain text HTML | stop scientific notation when saving to HTML with Aspose.Cells | C# HtmlSaveOptions ExportDataOptions.All example | apply text format @ to cells before HTML export | how to keep original column widths in Aspose.Cells HTML output
// Developer Intent: Save a workbook to HTML with numeric cells displayed as plain text rather than scientific notation.
// Use Cases: Web‑based financial reports that must show exact ID strings and precise decimal values. | Archiving spreadsheets on a website while preserving readable numeric formats. | Generating invoices or data tables where scientific notation would confuse end users.
// AI Prompts: Generate C# code using Aspose.Cells to export a worksheet to HTML with all numeric cells formatted as plain text, avoiding scientific notation. | Show how to set the Excel custom format "@" on specific cells before saving with HtmlSaveOptions. | Explain the role of HtmlExportDataOptions.All in ensuring cell values are included in the HTML output.

using System;
using Aspose.Cells;

namespace ExportNumericAsPlainText
{
    // Demonstrates how to create a workbook, apply the text format "@" to numeric cells, configure HtmlSaveOptions (ExportDataOptions.All) and save to HTML so large integers and tiny decimals are rendered as plain text instead of scientific notation.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add numeric values that would normally be shown in scientific notation
            cells["A1"].PutValue(123456789012345);
            cells["A2"].PutValue(0.0000001234);

            // Apply a text style to the cells so they are exported as plain text
            Style textStyle = workbook.CreateStyle();
            // "@" is the Excel format code for text
            textStyle.Custom = "@";
            cells["A1"].SetStyle(textStyle);
            cells["A2"].SetStyle(textStyle);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Ensure all data (including cell values) is exported
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.All;
            // Optional: keep the original column widths (set to false to ignore width truncation)
            htmlOptions.FormatDataIgnoreColumnWidth = false;

            // Save the workbook as HTML; numeric cells will appear as plain text
            workbook.Save("NumericPlainText.html", htmlOptions);
        }
    }
}
