// Title: C# – Export Excel to a Single HTML File with Inline CSS Using TableCssId (Aspose.Cells)
// Description: Demonstrates how to use Aspose.Cells for .NET to save a workbook as one HTML document with CSS embedded inline and class names prefixed by TableCssId, enabling instant preview without external style sheets.
// Keywords: Aspose.Cells HTML export | inline CSS Aspose.Cells | TableCssId prefix | SaveAsSingleFile .NET | C# export Excel to HTML | embedded CSS in HTML | self‑contained HTML workbook
// Common Searches: Aspose.Cells embed CSS in exported HTML | TableCssId usage in HtmlSaveOptions | C# save Excel as single HTML file | inline CSS for Aspose.Cells HTML output | how to prefix CSS classes when exporting Excel to HTML
// Developer Intent: Create a self‑contained HTML preview of an Excel workbook with custom‑prefixed CSS classes defined inline.
// Use Cases: Quickly display spreadsheet data on a web page without loading external CSS files. | Avoid class‑name collisions by applying a unique TableCssId prefix to all exported table styles. | Generate a single HTML file suitable for email bodies or documentation where external resources are prohibited.
// AI Prompts: Show how to change the background color in the CssStyles string while keeping the TableCssId prefix. | Provide C# code that reads an existing .xlsx file and exports it to a single HTML file with inline CSS using a custom TableCssId value. | Explain the interaction between SaveAsSingleFile and TableCssId in producing a self‑contained HTML preview.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells for .NET to save a workbook as one HTML document with CSS embedded inline and class names prefixed by TableCssId, enabling instant preview without external style sheets.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["A3"].PutValue("Alice");
        worksheet.Cells["B3"].PutValue(25);

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Prefix for CSS class names that will be applied to table elements
        saveOptions.TableCssId = "custom";

        // Save as a single HTML file so that CSS is embedded inline
        saveOptions.SaveAsSingleFile = true;

        // Define inline CSS using the specified prefix
        saveOptions.CssStyles = @"
            .custom-table { border-collapse: collapse; width: 100%; }
            .custom-tr:nth-child(even) { background-color: #f2f2f2; }
            .custom-td, .custom-th { border: 1px solid #ddd; padding: 8px; }
        ";

        // Export the workbook to HTML with the configured options
        workbook.Save("preview.html", saveOptions);
    }
}
