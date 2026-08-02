// Title: Export Aspose.Cells Workbook to HTML with Embedded CSS vs External CSS and Compare File Sizes (C#)
// Description: Shows how to save an Aspose.Cells workbook as a single HTML file with embedded CSS and as separate HTML/CSS files, then reads and prints the sizes of the generated files using C#.
// Keywords: Aspose.Cells | C# | HTML export | embedded CSS | external CSS | HtmlSaveOptions | ExportWorksheetCSSSeparately | single HTML file | file size comparison | performance optimization
// Common Searches: Aspose.Cells export workbook to HTML with embedded CSS | Aspose.Cells generate external CSS file when saving as HTML | compare HTML file size with embedded CSS versus external CSS Aspose | HtmlSaveOptions SaveAsSingleFile true example | ExportWorksheetCSSSeparately false C# sample
// Developer Intent: Produce HTML output from a workbook in two formats—embedded CSS and external CSS—and programmatically compare their file sizes to evaluate storage and bandwidth impact.
// Use Cases: Create a self‑contained HTML report that can be emailed or viewed offline without additional files. | Generate HTML with a linked stylesheet for web applications to enable browser caching and reduce HTML payload. | Measure the size difference between embedded and external CSS to guide performance‑related decisions for large spreadsheets. | Automate size reporting in CI/CD pipelines for document‑to‑HTML conversion processes.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a single HTML file with embedded CSS and returns the file size in bytes. | Show how to compress the generated HTML and CSS files using GZip after saving with Aspose.Cells. | Explain how to read the external CSS file produced by Aspose.Cells and inline its contents into the HTML for a custom combined output. | Create a PowerShell script that runs the compiled program and logs the size differences to a CSV file. | Generate a unit test that asserts the embedded HTML size is smaller than the combined size of external HTML plus CSS.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Shows how to save an Aspose.Cells workbook as a single HTML file with embedded CSS and as separate HTML/CSS files, then reads and prints the sizes of the generated files using C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some formatted data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data with different styles to generate CSS rules
            sheet.Cells["A1"].PutValue("Header");
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);

            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(123);
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.Color = System.Drawing.Color.Green;
            sheet.Cells["A2"].SetStyle(dataStyle);
            sheet.Cells["B2"].SetStyle(dataStyle);

            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(456);
            sheet.Cells["A3"].SetStyle(dataStyle);
            sheet.Cells["B3"].SetStyle(dataStyle);

            // -----------------------------------------------------------------
            // Export with embedded CSS (all CSS inside the HTML file)
            // -----------------------------------------------------------------
            HtmlSaveOptions embeddedOptions = new HtmlSaveOptions();
            embeddedOptions.ExportWorksheetCSSSeparately = false; // embed CSS
            embeddedOptions.SaveAsSingleFile = true; // single HTML file
            string embeddedHtmlPath = "EmbeddedCss.html";
            workbook.Save(embeddedHtmlPath, embeddedOptions);

            // -----------------------------------------------------------------
            // Export with external CSS (CSS written to a separate .css file)
            // -----------------------------------------------------------------
            HtmlSaveOptions externalOptions = new HtmlSaveOptions();
            externalOptions.ExportWorksheetCSSSeparately = true; // separate CSS file
            externalOptions.SaveAsSingleFile = false; // default behavior
            string externalHtmlPath = "ExternalCss.html";
            workbook.Save(externalHtmlPath, externalOptions);
            // The CSS file will be generated alongside the HTML (e.g., sheet0.css)

            // -----------------------------------------------------------------
            // Compare file sizes
            // -----------------------------------------------------------------
            long embeddedSize = new FileInfo(embeddedHtmlPath).Length;
            long externalSize = new FileInfo(externalHtmlPath).Length;

            Console.WriteLine($"Embedded HTML size: {embeddedSize} bytes");
            Console.WriteLine($"External CSS HTML size: {externalSize} bytes");

            // If needed, also display the size of the generated CSS file
            string cssFilePath = Path.Combine(Path.GetDirectoryName(externalHtmlPath) ?? "", "sheet0.css");
            if (File.Exists(cssFilePath))
            {
                long cssSize = new FileInfo(cssFilePath).Length;
                Console.WriteLine($"External CSS file size (sheet0.css): {cssSize} bytes");
            }
        }
    }
}
