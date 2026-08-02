// Title: Export Excel to HTML with Exact Column Widths using Aspose.Cells for .NET
// Description: Loads an Excel workbook, optionally sets column widths, configures HtmlSaveOptions (WidthScalable = false, FormatDataIgnoreColumnWidth = false) and saves the file as HTML where each column is rendered with a fixed pixel width via CSS.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | preserve column width | HtmlSaveOptions | WidthScalable false | FormatDataIgnoreColumnWidth | CSS width export
// Common Searches: Aspose.Cells keep column width when exporting to HTML | HTML export with fixed column sizes using Aspose.Cells .NET | How to retain Excel column widths in generated HTML | Aspose.Cells HtmlSaveOptions column width settings
// Developer Intent: Generate an HTML representation of an Excel workbook that maintains the original column widths using CSS width attributes.
// Use Cases: Create web‑ready reports that mirror the layout of the source spreadsheet. | Provide an accurate HTML preview of Excel data in a web application. | Produce email‑compatible HTML tables where column dimensions must stay consistent.
// AI Prompts: Show how to export column widths as percentages instead of pixels with Aspose.Cells. | Give a code sample that writes the HTML to a MemoryStream and returns it from an ASP.NET Core controller while preserving widths. | Explain how to attach a custom CSS class to the generated table without affecting column‑width preservation.

using System;
using Aspose.Cells;

// Loads an Excel workbook, optionally sets column widths, configures HtmlSaveOptions (WidthScalable = false, FormatDataIgnoreColumnWidth = false) and saves the file as HTML where each column is rendered with a fixed pixel width via CSS.
class Program
{
    static void Main()
    {
        // Load an existing Excel workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Example: set custom column widths to demonstrate preservation
        // (optional – remove if widths are already defined in the source file)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.SetColumnWidth(0, 20); // Column A width
        sheet.Cells.SetColumnWidth(1, 30); // Column B width

        // Configure HTML save options to keep original column widths
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Export column widths as fixed pixel values (CSS width) – default behavior,
        // but set explicitly for clarity.
        htmlOptions.WidthScalable = false;

        // Ensure cell content respects column width (do not ignore it).
        htmlOptions.FormatDataIgnoreColumnWidth = false;

        // Save the workbook as an HTML file with the above options.
        workbook.Save("output.html", htmlOptions);
    }
}
