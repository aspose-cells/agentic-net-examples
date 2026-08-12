// Title: Export Excel to HTML with Inline CSS and CSS Custom Properties using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, applies bold and italic formatting, disables external CSS, enables CSS custom properties, and saves the file as HTML containing only inline styles and reusable custom property definitions.
// Keywords: Aspose.Cells HTML export | inline CSS Aspose.Cells | DisableCss true | EnableCssCustomProperties | C# Excel to HTML | Aspose.Cells HtmlSaveOptions | .NET workbook to HTML | custom CSS properties Excel export
// Common Searches: Aspose.Cells export HTML with inline styles | How to disable external CSS when saving Excel as HTML in .NET | Enable CSS custom properties in Aspose.Cells HtmlSaveOptions | C# code sample for Aspose.Cells HTML export without stylesheet | Generate lightweight HTML from Excel using Aspose.Cells
// Developer Intent: Create an HTML file from an Excel workbook that embeds all styling inline and leverages CSS custom properties for compact, stylesheet‑free output.
// Use Cases: Email‑compatible HTML reports where external CSS files are blocked. | Embedding Excel‑derived tables directly into web pages without additional stylesheet assets. | Reducing page weight by reusing style values through CSS custom properties in generated HTML.
// AI Prompts: Explain the impact of DisableCss and EnableCssCustomProperties on the HTML produced by Aspose.Cells. | Modify the sample to embed worksheet images as base64 data URIs during HTML export. | Write a unit test that verifies the exported HTML contains inline style attributes and CSS custom property definitions.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // C# example that creates a workbook, applies bold and italic formatting, disables external CSS, enables CSS custom properties, and saves the file as HTML containing only inline styles and reusable custom property definitions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with some formatting
            sheet.Cells["A1"].PutValue("Hello");
            Style styleA1 = sheet.Cells["A1"].GetStyle();
            styleA1.Font.IsBold = true;
            styleA1.Font.Color = System.Drawing.Color.Blue;
            sheet.Cells["A1"].SetStyle(styleA1);

            sheet.Cells["B2"].PutValue("World");
            Style styleB2 = sheet.Cells["B2"].GetStyle();
            styleB2.Font.IsItalic = true;
            styleB2.Font.Color = System.Drawing.Color.Green;
            sheet.Cells["B2"].SetStyle(styleB2);

            // Configure HTML save options:
            // - DisableCss = true  => only inline styles, no external CSS files
            // - EnableCssCustomProperties = true => use CSS custom properties for optimization
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableCss = true;
            htmlOptions.EnableCssCustomProperties = true;

            // Save the workbook as HTML with the specified options
            string outputPath = "ExportedWithInlineAndCustomProps.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with inline styles and CSS custom properties enabled.");
        }
    }
}
