// Title: Export Excel to HTML with Gridlines and TableCssId Prefix Using Aspose.Cells for .NET
// Description: Shows how to make worksheet gridlines visible, apply thin borders, and configure HtmlSaveOptions (ExportGridLines, TableCssId, ExportSimilarBorderStyle) to produce an HTML file that retains Excel cell borders and prefixes CSS selectors, with a complete C# example.
// Keywords: Aspose.Cells | C# | .NET | ExportGridLines | HtmlSaveOptions | TableCssId | ExportSimilarBorderStyle | gridlines HTML export | Excel to HTML | preserve cell borders | web report generation | GitHub | US | Europe
// Common Searches: Aspose.Cells export gridlines to HTML | How to keep Excel gridlines when saving as HTML | TableCssId usage in Aspose.Cells HTML export | Enable ExportSimilarBorderStyle in Aspose.Cells | C# export worksheet to HTML with borders | Aspose.Cells HtmlSaveOptions sample code
// Developer Intent: The developer needs to export an Excel worksheet to HTML while preserving gridlines and ensuring CSS selectors are correctly prefixed for border styling.
// Use Cases: Create HTML reports that match the visual layout of the original Excel file. | Generate web‑ready tables with prefixed CSS classes to avoid style conflicts. | Export workbooks that contain thin or custom borders, using ExportSimilarBorderStyle to approximate unsupported styles.
// AI Prompts: Show how to disable ExportGridLines but keep custom borders in the HTML output. | Provide an example of linking an external CSS file to style a table exported with TableCssId. | Explain the impact of ExportSimilarBorderStyle on complex border rendering in Aspose.Cells HTML export.

using System;
using Aspose.Cells;

// Shows how to make worksheet gridlines visible, apply thin borders, and configure HtmlSaveOptions (ExportGridLines, TableCssId, ExportSimilarBorderStyle) to produce an HTML file that retains Excel cell borders and prefixes CSS selectors, with a complete C# example.
class ExportWithGridlines
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Make gridlines visible in the worksheet
            worksheet.IsGridlinesVisible = true;

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue("Item 1");
            worksheet.Cells["B2"].PutValue("Item 2");

            // Apply a thin border to the range to demonstrate border export
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Use CreateRange for a multi‑cell range
            worksheet.Cells.CreateRange("A1:B2").SetStyle(borderStyle);

            // Configure HTML save options:
            // - ExportGridLines = true to retain cell borders as gridlines
            // - TableCssId provides a prefix for CSS selectors (e.g., tr.my-table, td.my-table)
            // - ExportSimilarBorderStyle ensures unsupported border styles are approximated
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = true,
                TableCssId = "my-table",
                ExportSimilarBorderStyle = true
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
