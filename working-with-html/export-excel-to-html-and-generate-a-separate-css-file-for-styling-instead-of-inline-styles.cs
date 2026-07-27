// Title: Export Excel to HTML with an external CSS file using Aspose.Cells C#
// Description: Shows how to build a workbook, apply header styling, and save it as HTML while generating a separate .css stylesheet via HtmlSaveOptions.ExportWorksheetCSSSeparately in C#.
// Keywords: Aspose.Cells | C# | HTML export | external CSS | ExportWorksheetCSSSeparately | Excel to HTML | separate stylesheet | Aspose HtmlSaveOptions
// Common Searches: Aspose.Cells export HTML external CSS C# | HtmlSaveOptions ExportWorksheetCSSSeparately example | Save Excel as HTML without inline styles Aspose | C# generate separate CSS file from workbook | How to create external stylesheet when converting Excel to HTML
// Developer Intent: Create an HTML version of an Excel workbook where all visual formatting is placed in an external CSS file rather than inline styles.
// Use Cases: Publish spreadsheet data on a website with a shared stylesheet for consistent branding. | Automate report pipelines that deliver HTML pages and a separate CSS file to a content management system. | Integrate Excel‑derived tables into web applications while keeping presentation logic in maintainable CSS files.
// AI Prompts: Generate code that customizes the filename and folder of the CSS file produced by HtmlSaveOptions. | Provide a snippet to read the generated CSS, add a custom class, and reference it in the HTML output. | Explain how to disable all inline styling so the HTML relies exclusively on the external stylesheet when converting a workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to build a workbook, apply header styling, and save it as HTML while generating a separate .css stylesheet via HtmlSaveOptions.ExportWorksheetCSSSeparately in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Apply simple formatting to demonstrate CSS generation
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Configure HTML save options to export CSS to a separate file
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportWorksheetCSSSeparately = true; // Generates a .css file instead of inline styles

            // Define output directory and file names
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AsposeHtmlExport");
            Directory.CreateDirectory(outputDir);
            string htmlPath = Path.Combine(outputDir, "Workbook.html");

            // Save the workbook as HTML; a separate CSS file (Workbook.css) will be created in the same folder
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {htmlPath}");
            Console.WriteLine($"Separate CSS file generated alongside the HTML.");
        }
    }
}
