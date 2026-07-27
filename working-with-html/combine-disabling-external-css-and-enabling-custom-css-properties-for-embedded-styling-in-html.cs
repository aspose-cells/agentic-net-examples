// Title: Aspose.Cells C# – Export Workbook to HTML with Inline Styles and CSS Custom Properties
// Description: Demonstrates how to save an Aspose.Cells workbook as HTML using HtmlSaveOptions with DisableCss=true (inline styling only) and EnableCssCustomProperties=true (reusable CSS variables). The example applies bold/blue and italic/green formatting, then generates a self‑contained HTML file.
// Keywords: Aspose.Cells HTML export | C# HtmlSaveOptions DisableCss | EnableCssCustomProperties | inline CSS Aspose.Cells | custom CSS variables workbook | export workbook to HTML without external CSS
// Common Searches: Aspose.Cells disable external CSS when saving to HTML | How to enable CSS custom properties in Aspose.Cells HTML output | C# export Excel to HTML with only inline styles | HtmlSaveOptions DisableCss true example | Aspose.Cells HTML custom properties for repeated colors
// Developer Intent: Create an HTML representation of a workbook that contains only inline style attributes and leverages CSS custom properties to reduce redundancy.
// Use Cases: Generate HTML reports for email or intranet portals where linked CSS files are prohibited. | Minimize HTML payload by reusing colors, fonts, and other styles through CSS variables. | Combine inline cell styling with a small global CSS block when SaveAsSingleFile is enabled.
// AI Prompts: Show the C# code to export an Aspose.Cells workbook to HTML with DisableCss and EnableCssCustomProperties enabled. | Explain how CSS custom properties affect the markup produced by HtmlSaveOptions. | Provide a snippet that adds a custom CSS block while using SaveAsSingleFile in Aspose.Cells HTML export.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlStylingDemo
{
    // Demonstrates how to save an Aspose.Cells workbook as HTML using HtmlSaveOptions with DisableCss=true (inline styling only) and EnableCssCustomProperties=true (reusable CSS variables). The example applies bold/blue and italic/green formatting, then generates a self‑contained HTML file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data with formatting
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A1"].GetStyle().Font.IsBold = true;
            sheet.Cells["A1"].GetStyle().Font.Color = System.Drawing.Color.Blue;

            sheet.Cells["A2"].PutValue("Data");
            sheet.Cells["A2"].GetStyle().Font.IsItalic = true;
            sheet.Cells["A2"].GetStyle().Font.Color = System.Drawing.Color.Green;

            // Configure HTML save options:
            // - DisableCss = true  => only inline styles, no external CSS files.
            // - EnableCssCustomProperties = true => use CSS custom properties for repeated resources.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableCss = true;
            htmlOptions.EnableCssCustomProperties = true;

            // Optional: add a small custom CSS block (works when SaveAsSingleFile is true)
            // Here we keep it simple; the main goal is to demonstrate the two flags.
            // htmlOptions.SaveAsSingleFile = true;
            // htmlOptions.CssStyles = "body { font-family: Arial; }";

            // Save the workbook as HTML with the configured options
            workbook.Save("StyledOutput.html", htmlOptions);

            Console.WriteLine("HTML file generated with inline styles and CSS custom properties enabled.");
        }
    }
}
