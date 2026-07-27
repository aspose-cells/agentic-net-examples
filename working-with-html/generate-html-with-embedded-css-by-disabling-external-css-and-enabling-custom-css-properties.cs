// Title: Export Excel to HTML with Inline Styles & CSS Custom Properties using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, apply bold red and italic blue formatting, configure HtmlSaveOptions to disable external CSS, enable CSS custom properties, add a custom CSS block, and save the file as HTML that relies on inline styles and reusable custom properties.
// Keywords: Aspose.Cells HTML export | inline CSS Aspose.Cells | Disable external CSS .NET | EnableCssCustomProperties | custom CSS block Aspose.Cells | Excel to HTML C# | Aspose.Cells HtmlSaveOptions | single‑file HTML export | web‑ready Excel report
// Common Searches: Aspose.Cells export Excel to HTML with inline styles | How to disable external CSS in Aspose.Cells HTML output | Enable CSS custom properties when saving workbook as HTML C# | Add custom CSS block to Aspose.Cells generated HTML | Save Excel as single‑file HTML using Aspose.Cells
// Developer Intent: The developer wants to generate an HTML version of an Excel workbook that contains only inline styling and leverages CSS custom properties, optionally embedding a small custom stylesheet.
// Use Cases: Create web‑ready reports from Excel without external stylesheet dependencies. | Reduce duplicated style definitions by using CSS custom properties for repeated cell formats. | Inject a site‑specific CSS snippet into the exported HTML to preserve branding while keeping the output minimal.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a single‑file HTML with embedded CSS and CSS custom properties enabled. | Show how to set HtmlSaveOptions.DisableCss = true, EnableCssCustomProperties = true, and add a custom CSS block for Aspose.Cells HTML export. | Explain the impact of EnableCssCustomProperties on the generated HTML and how inline styles reference those custom properties.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Demonstrates how to create a workbook, apply bold red and italic blue formatting, configure HtmlSaveOptions to disable external CSS, enable CSS custom properties, add a custom CSS block, and save the file as HTML that relies on inline styles and reusable custom properties.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data with various formatting
            sheet.Cells["A1"].PutValue("Bold Red Text");
            Style styleA1 = sheet.Cells["A1"].GetStyle();
            styleA1.Font.IsBold = true;
            styleA1.Font.Color = System.Drawing.Color.Red;
            sheet.Cells["A1"].SetStyle(styleA1);

            sheet.Cells["B2"].PutValue("Italic Blue Text");
            Style styleB2 = sheet.Cells["B2"].GetStyle();
            styleB2.Font.IsItalic = true;
            styleB2.Font.Color = System.Drawing.Color.Blue;
            sheet.Cells["B2"].SetStyle(styleB2);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable external CSS (use only inline styles)
            htmlOptions.DisableCss = true;

            // Enable CSS custom properties to optimize repeated resources
            htmlOptions.EnableCssCustomProperties = true;

            // Optional: add a small custom CSS block (will be embedded if SaveAsSingleFile is true)
            // Here we keep SaveAsSingleFile false, so this CSS will be written to a separate file.
            // If you want it embedded, set htmlOptions.SaveAsSingleFile = true;
            htmlOptions.CssStyles = @"
                body { font-family: Arial, sans-serif; }
                .highlight { background-color: #ffff99; }
            ";

            // Save the workbook as HTML with the configured options
            string outputPath = "ExportedWithInlineAndCustomProperties.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with inline styles and CSS custom properties enabled.");
        }
    }
}
