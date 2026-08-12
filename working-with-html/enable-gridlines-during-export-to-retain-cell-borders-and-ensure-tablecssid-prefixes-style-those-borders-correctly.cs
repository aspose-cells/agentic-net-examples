// Title: Export Excel to HTML with Gridlines and Custom TableCssId using Aspose.Cells for .NET
// Description: Demonstrates how to enable worksheet gridlines, apply thin borders, and configure HtmlSaveOptions (ExportGridLines, TableCssId, ExportSimilarBorderStyle) so the generated HTML retains cell borders and can be styled via a custom CSS ID.
// Keywords: Aspose.Cells HTML export | ExportGridLines .NET | TableCssId styling | preserve Excel borders HTML | C# Aspose.Cells example | gridlines to HTML | ExportSimilarBorderStyle | custom CSS table ID
// Common Searches: Aspose.Cells keep gridlines when saving as HTML | How to use TableCssId in HtmlSaveOptions | Export Excel borders to HTML with Aspose.Cells | Enable ExportSimilarBorderStyle for HTML output | C# export workbook to HTML with custom CSS ID
// Developer Intent: Generate an HTML file from an Excel workbook that keeps gridlines and cell borders and assigns a custom CSS ID to the table for styling.
// Use Cases: Create web‑ready reports that visually match the original Excel layout. | Integrate exported tables into existing site themes by targeting a specific CSS ID. | Ensure consistent border rendering across browsers with ExportSimilarBorderStyle.
// AI Prompts: Show how to change TableCssId to 'report-table' while preserving gridlines and borders. | Explain the effect of ExportSimilarBorderStyle on border appearance in Chrome, Firefox, and Edge. | Provide CSS rules that style the table with ID 'custom-table' to highlight cell borders after export.

using System;
using Aspose.Cells;

// Demonstrates how to enable worksheet gridlines, apply thin borders, and configure HtmlSaveOptions (ExportGridLines, TableCssId, ExportSimilarBorderStyle) so the generated HTML retains cell borders and can be styled via a custom CSS ID.
class ExportWithGridlines
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Make gridlines visible in the worksheet
            sheet.IsGridlinesVisible = true;

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue("Data 2");

            // Apply a thin border to the range to demonstrate border export
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the style to the range A1:B2
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:B2");
            StyleFlag flag = new StyleFlag { All = true };
            range.ApplyStyle(borderStyle, flag);

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = true,               // Export gridlines so cell borders are retained
                TableCssId = "custom-table",          // Prefix for CSS selectors (e.g., tr, td) within the table
                ExportSimilarBorderStyle = true       // Use similar border style when browser does not support exact style
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("Exported.html", options);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
