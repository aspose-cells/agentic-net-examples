// Title: C# – Export Excel to HTML with Aspose.Cells, disable embedded CSS and link an external stylesheet
// Description: Demonstrates how to save a workbook as HTML using Aspose.Cells with HtmlSaveOptions.DisableCss, generate a custom CSS file, and programmatically insert a <link> tag into the exported HTML so styling is controlled by an external stylesheet.
// Keywords: Aspose.Cells HTML export C# | DisableCss option | external stylesheet link | post‑process HTML output | custom CSS for Excel reports
// Common Searches: Aspose.Cells disable CSS when saving as HTML | add external CSS to Aspose.Cells generated HTML | C# insert <link> tag into exported HTML | HtmlSaveOptions.DisableCss example
// Developer Intent: Create an HTML file from a workbook without embedded CSS and attach a custom external stylesheet for centralized styling.
// Use Cases: Generate HTML reports from Excel where design is managed via a site‑wide CSS file. | Automate post‑processing of Aspose.Cells HTML output to inject a stylesheet reference for branding consistency. | Produce reusable web‑ready tables that inherit styles from a shared stylesheet rather than inline definitions.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to HTML with no embedded CSS and then adds a <link> tag for an external stylesheet. | Explain the effect of HtmlSaveOptions.DisableCss and show how to programmatically insert a stylesheet reference into the generated HTML file. | Provide step‑by‑step instructions for creating a custom CSS file and linking it to HTML exported from Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to save a workbook as HTML using Aspose.Cells with HtmlSaveOptions.DisableCss, generate a custom CSS file, and programmatically insert a <link> tag into the exported HTML so styling is controlled by an external stylesheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Configure HTML save options to disable embedded CSS (use inline styles only)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableCss = true; // No external or embedded CSS will be generated

            // Define output paths
            string htmlPath = "output.html";
            string cssPath = "custom.css";

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);

            // Create an external CSS file with custom styling
            string customCss = @"
                body {
                    font-family: Arial, sans-serif;
                    background-color: #f9f9f9;
                }
                table {
                    border-collapse: collapse;
                    width: 100%;
                }
                td, th {
                    border: 1px solid #ccc;
                    padding: 8px;
                }
                th {
                    background-color: #e0e0e0;
                }";
            File.WriteAllText(cssPath, customCss);

            // Read the generated HTML, insert a <link> tag for the external stylesheet, and save it back
            string htmlContent = File.ReadAllText(htmlPath);
            // Find the closing </head> tag to insert the link before it
            int headCloseIndex = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
            if (headCloseIndex != -1)
            {
                string linkTag = $@"<link rel=""stylesheet"" type=""text/css"" href=""{cssPath}"">" + Environment.NewLine;
                htmlContent = htmlContent.Insert(headCloseIndex, linkTag);
                File.WriteAllText(htmlPath, htmlContent);
            }

            Console.WriteLine("HTML exported with inline styles and linked external CSS.");
        }
    }
}
