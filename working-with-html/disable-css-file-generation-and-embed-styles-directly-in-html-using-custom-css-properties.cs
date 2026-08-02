// Title: Export Aspose.Cells Workbook to Self‑Contained HTML with Inline CSS Using Custom Properties (C#)
// Description: Demonstrates how to save an Aspose.Cells workbook as a single HTML file with all styling embedded inline. The example sets HtmlSaveOptions.DisableCss to true and EnableCssCustomProperties to true, eliminating external CSS files and using CSS custom properties for compact, self‑contained output.
// Keywords: Aspose.Cells HtmlSaveOptions DisableCss | EnableCssCustomProperties C# | inline CSS Aspose.Cells HTML export | self‑contained HTML workbook | embed styles in HTML Aspose | no external CSS Aspose.Cells | custom CSS variables Excel to HTML
// Common Searches: Aspose.Cells export to HTML without external CSS | How to embed CSS custom properties in Aspose.Cells HTML output | C# save workbook as single HTML file with inline styles | DisableCss option Aspose.Cells example | EnableCssCustomProperties usage in Aspose.Cells
// Developer Intent: Export an Excel workbook to HTML where all formatting is included directly in the page, avoiding separate CSS files.
// Use Cases: Email‑ready reports that must not reference external style sheets. | Single‑file documentation portals where extra HTTP requests are undesirable. | Embedding formatted spreadsheet data into web pages that need instant rendering.
// AI Prompts: Show how to embed worksheet images as base64 strings in the generated HTML using Aspose.Cells. | Provide a snippet that defines a CSS variable for cell background color and applies it via custom properties. | Explain how to switch back to external CSS files while still using custom properties for selected styles.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to save an Aspose.Cells workbook as a single HTML file with all styling embedded inline. The example sets HtmlSaveOptions.DisableCss to true and EnableCssCustomProperties to true, eliminating external CSS files and using CSS custom properties for compact, self‑contained output.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data with some formatting
        worksheet.Cells["A1"].PutValue("Hello, Aspose!");
        Style cellStyle = worksheet.Cells["A1"].GetStyle();
        cellStyle.Font.IsBold = true;
        cellStyle.Font.Color = Color.Blue;
        worksheet.Cells["A1"].SetStyle(cellStyle);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Disable external CSS files; use only inline styles
        htmlOptions.DisableCss = true;

        // Enable CSS custom properties to embed styles directly in the HTML
        htmlOptions.EnableCssCustomProperties = true;

        // Save the workbook as HTML (lifecycle save)
        workbook.Save("output.html", htmlOptions);
    }
}
