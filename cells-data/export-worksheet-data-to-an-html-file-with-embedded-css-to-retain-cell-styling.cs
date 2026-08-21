// Title: Export Aspose.Cells Worksheet to HTML with Embedded CSS (C# .NET)
// Description: Creates a workbook, adds a product‑price table, applies bold white text on a dark‑blue header, and saves it as an HTML file using HtmlSaveOptions (ExportWorksheetCSSSeparately = false, DisableCss = false, ExcludeUnusedStyles = false) so all cell styles are embedded in the output.
// Keywords: Aspose.Cells | HTML export | embedded CSS | preserve cell styling | C# .NET | HtmlSaveOptions | ExportWorksheetCSSSeparately | DisableCss | ExcludeUnusedStyles | Windows desktop
// Common Searches: Aspose.Cells export worksheet to HTML with CSS | C# save Excel as styled HTML Aspose | HtmlSaveOptions embed CSS example | retain cell formatting when converting to HTML | Aspose.Cells HTML output with embedded styles
// Developer Intent: Generate an HTML document from a workbook that contains all formatting information as embedded CSS, eliminating external style files.
// Use Cases: Publish a styled price list on a website without separate CSS files. | Provide an on‑the‑fly HTML preview of an Excel report in a .NET application. | Create email‑ready HTML attachments that keep the original spreadsheet appearance.
// AI Prompts: Write C# code using Aspose.Cells to export a worksheet to HTML with embedded CSS, keeping header font weight and background color. | Show how to configure HtmlSaveOptions (ExportWorksheetCSSSeparately, DisableCss, ExcludeUnusedStyles) for full style retention in HTML output. | Provide a complete example that builds a small data table, applies custom styling, and saves it as an HTML file with embedded CSS on the user's desktop.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Creates a workbook, adds a product‑price table, applies bold white text on a dark‑blue header, and saves it as an HTML file using HtmlSaveOptions (ExportWorksheetCSSSeparately = false, DisableCss = false, ExcludeUnusedStyles = false) so all cell styles are embedded in the output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Apply some styling to demonstrate retained CSS
            Style headerStyle = sheet.Cells["A1"].GetStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Configure HTML save options to embed CSS (default behavior)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Ensure CSS is embedded in the HTML file
                ExportWorksheetCSSSeparately = false,
                // Use CSS (not only inline styles)
                DisableCss = false,
                // Keep all styles (optional, ensures no style is omitted)
                ExcludeUnusedStyles = false
            };

            // Define output path
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "WorkbookExport.html");

            // Save the workbook as HTML with embedded CSS
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file with embedded CSS saved to: {outputPath}");
        }
    }
}
