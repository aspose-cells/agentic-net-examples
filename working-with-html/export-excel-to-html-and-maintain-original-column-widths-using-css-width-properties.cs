// Title: Export Excel to HTML with Exact Column Widths using Aspose.Cells for .NET
// Description: Demonstrates how to save a workbook as HTML while preserving the original column widths. The example sets column widths in characters, configures HtmlSaveOptions (FormatDataIgnoreColumnWidth = false, WidthScalable = false), and generates CSS width rules that match the Excel layout.
// Keywords: Aspose.Cells HTML export | C# Excel to HTML | preserve column width | Css width from Excel | HtmlSaveOptions FormatDataIgnoreColumnWidth | WidthScalable false | fixed pixel column size | .NET spreadsheet conversion | Excel column width CSS | export workbook as HTML
// Common Searches: Aspose.Cells keep column widths when exporting to HTML | C# export Excel to HTML with fixed column sizes | HtmlSaveOptions column width settings | How to generate CSS width for Excel columns using Aspose | Export Excel worksheet to HTML preserving layout
// Developer Intent: Export an Excel worksheet to HTML and retain the exact column widths through CSS styling.
// Use Cases: Create web‑ready reports that mirror the spreadsheet’s column layout. | Build dashboards where table columns must stay aligned across browsers. | Convert Excel templates into static HTML pages for newsletters or documentation while keeping the original column dimensions.
// AI Prompts: Show how to output column widths as percentage values instead of fixed pixels with Aspose.Cells. | Add a custom CSS class to the generated HTML table while preserving column widths. | Export multiple worksheets to separate HTML files, each maintaining its own column width settings.

using System;
using Aspose.Cells;

namespace ExportExcelToHtmlWithColumnWidths
{
    // Demonstrates how to save a workbook as HTML while preserving the original column widths. The example sets column widths in characters, configures HtmlSaveOptions (FormatDataIgnoreColumnWidth = false, WidthScalable = false), and generates CSS width rules that match the Excel layout.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path) or create a new one.
            Workbook workbook = new Workbook(); // new workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data.
            sheet.Cells["A1"].PutValue("Short");
            sheet.Cells["B1"].PutValue("A much longer text that should respect column width");
            sheet.Cells["C1"].PutValue(12345);

            // Set column widths (in characters). These widths will be exported as CSS width properties.
            sheet.Cells.SetColumnWidth(0, 12); // Column A
            sheet.Cells.SetColumnWidth(1, 30); // Column B
            sheet.Cells.SetColumnWidth(2, 15); // Column C

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Ensure column widths are not ignored (default is false, set explicitly for clarity).
            htmlOptions.FormatDataIgnoreColumnWidth = false;

            // Export column widths as fixed pixel values (not scalable). This keeps the original widths.
            htmlOptions.WidthScalable = false;

            // Save the workbook as HTML. The generated HTML will contain CSS rules that preserve the column widths.
            workbook.Save("ExportedWithColumnWidths.html", htmlOptions);
        }
    }
}
