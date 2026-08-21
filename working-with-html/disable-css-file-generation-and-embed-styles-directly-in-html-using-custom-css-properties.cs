// Title: Aspose.Cells for .NET – Export Workbook to a Single HTML File with Embedded CSS and Custom Properties
// Description: Shows how to configure HtmlSaveOptions in C# to generate a self‑contained HTML file from an Aspose.Cells workbook. The sample disables external CSS files, embeds all styles inside a <style> tag, and enables CSS custom properties for flexible theming.
// Keywords: Aspose.Cells HTML export | C# embed CSS in HTML | HtmlSaveOptions SaveAsSingleFile | EnableCssCustomProperties | single file HTML Aspose.Cells | disable external stylesheet | embedded CSS Aspose.Cells .NET
// Common Searches: Aspose.Cells export workbook to HTML with embedded CSS | How to disable external CSS files in Aspose.Cells HTML export | Enable CSS custom properties in Aspose.Cells C# | Save Aspose.Cells workbook as single HTML file | Aspose.Cells HtmlSaveOptions CSS settings
// Developer Intent: Create a single HTML document from a workbook where all styling is included in the file itself, using CSS custom properties and without generating separate stylesheet files.
// Use Cases: Email‑compatible HTML reports where external stylesheets are blocked. | Offline‑viewable HTML documents that retain spreadsheet formatting without extra files. | Web pages that rely on CSS variables for dynamic theming while keeping the markup portable.
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook as one HTML file with all CSS embedded and CSS custom properties enabled. | Explain the impact of EnableCssCustomProperties on the HTML output produced by Aspose.Cells. | Provide a step‑by‑step guide to disable external CSS generation and embed styles directly using HtmlSaveOptions in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to configure HtmlSaveOptions in C# to generate a self‑contained HTML file from an Aspose.Cells workbook. The sample disables external CSS files, embeds all styles inside a <style> tag, and enables CSS custom properties for flexible theming.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data with formatting
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Apply bold font to header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.Blue;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Embed all CSS into the generated HTML (no external CSS files)
            htmlOptions.SaveAsSingleFile = true;          // merges CSS into <style> tag
            htmlOptions.DisableCss = false;              // allow CSS (inline or <style>)
            htmlOptions.EnableCssCustomProperties = true; // use CSS custom properties for optimization

            // Optional: add additional CSS if needed
            htmlOptions.CssStyles = @"
                body { font-family: Arial, sans-serif; margin: 20px; }
                table { border-collapse: collapse; width: 50%; }
                td, th { border: 1px solid #ddd; padding: 8px; }
                th { background-color: var(--header-bg, #f2f2f2); }";

            // Save the workbook as a single HTML file with embedded styles
            string outputPath = "WorkbookWithEmbeddedCss.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
