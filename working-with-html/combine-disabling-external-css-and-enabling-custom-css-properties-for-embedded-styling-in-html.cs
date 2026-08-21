// Title: Export Excel to a single self‑contained HTML file with embedded CSS and CSS custom properties using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, applies header styling, and uses HtmlSaveOptions to embed all CSS (SaveAsSingleFile = true), enable CSS custom properties, add custom CSS rules, and save the result as one HTML file.
// Keywords: Aspose.Cells HTML export | C# embed CSS | SaveAsSingleFile | EnableCssCustomProperties | custom CSS styles | self‑contained HTML report | Aspose.Cells .NET
// Common Searches: Aspose.Cells embed CSS in HTML export | Save workbook as single HTML file C# | Enable CSS custom properties Aspose.Cells | Add custom CSS rules with HtmlSaveOptions | Disable external CSS files Aspose.Cells
// Developer Intent: Generate a single HTML file from an Excel workbook with all styling embedded and CSS custom properties activated.
// Use Cases: Send a complete HTML report via email without external style sheets. | Reduce HTML size by reusing base64 images and repeated styles through CSS custom properties. | Apply project‑specific visual tweaks (fonts, borders, spacing) directly during export.
// AI Prompts: Show C# code to export an Aspose.Cells workbook to HTML with SaveAsSingleFile and EnableCssCustomProperties enabled. | How can I add custom CSS rules via HtmlSaveOptions.CssStyles while preventing external CSS files? | Explain the benefits of EnableCssCustomProperties for embedded resources like base64 images in Aspose.Cells HTML output.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Creates a workbook, applies header styling, and uses HtmlSaveOptions to embed all CSS (SaveAsSingleFile = true), enable CSS custom properties, add custom CSS rules, and save the result as one HTML file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data with formatting to demonstrate CSS usage
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Apply bold font to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable external CSS files (embed all CSS into the HTML)
            // By saving as a single file, no separate CSS files are generated.
            htmlOptions.SaveAsSingleFile = true;

            // Enable CSS custom properties to optimize repeated resources (e.g., base64 images)
            htmlOptions.EnableCssCustomProperties = true;

            // Optional: add additional custom CSS rules
            htmlOptions.CssStyles = @"
                body { font-family: Arial, sans-serif; margin: 20px; }
                table { border-collapse: collapse; width: 100%; }
                td, th { border: 1px solid #ddd; padding: 8px; }
            ";

            // Save the workbook as HTML with the configured options
            workbook.Save("ExportedWorkbook.html", htmlOptions);

            Console.WriteLine("HTML file generated with embedded CSS and custom properties.");
        }
    }
}
