// Title: Aspose.Cells C# – Export to HTML without embedded CSS and link an external stylesheet
// Description: Learn how to save an Aspose.Cells workbook as HTML with HtmlSaveOptions.DisableCss, create a custom CSS file, and programmatically insert a <link> tag so the page uses external styling.
// Keywords: Aspose.Cells HTML export C# | DisableCss Aspose.Cells | external stylesheet Aspose.Cells | link CSS after HTML save | custom CSS Aspose.Cells HTML output
// Common Searches: Aspose.Cells disable embedded CSS C# | Add external CSS to Aspose.Cells HTML export | HtmlSaveOptions.DisableCss example | Insert stylesheet link into generated HTML Aspose.Cells | Post‑process Aspose.Cells HTML for custom styling
// Developer Intent: Create an HTML file from a workbook without embedded CSS and attach a custom external stylesheet automatically.
// Use Cases: Generate HTML reports that share a corporate stylesheet for consistent branding. | Batch‑convert many Excel files to HTML while reusing a single CSS file to simplify maintenance. | Automate post‑processing of Aspose.Cells HTML output to replace inline styles with a shared stylesheet.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with embedded CSS disabled and then add a <link> tag for an external CSS file. | Provide a method that receives the HTML string from Aspose.Cells and injects a stylesheet link before </head>, handling missing </head> gracefully. | Explain the effect of HtmlSaveOptions.DisableCss on the generated HTML and how to combine it with a custom external stylesheet for uniform styling.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Learn how to save an Aspose.Cells workbook as HTML with HtmlSaveOptions.DisableCss, create a custom CSS file, and programmatically insert a <link> tag so the page uses external styling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
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
            htmlOptions.DisableCss = true; // disables generation of separate CSS files

            // Define output file paths
            string htmlFilePath = "output.html";
            string cssFilePath = "styles.css";

            // Save the workbook as HTML with the specified options
            workbook.Save(htmlFilePath, htmlOptions);

            // Create an external CSS file with custom styling
            string customCss = @"
                body {
                    font-family: Arial, Helvetica, sans-serif;
                    background-color: #f9f9f9;
                    margin: 20px;
                }
                table {
                    border-collapse: collapse;
                    width: 100%;
                }
                td, th {
                    border: 1px solid #ddd;
                    padding: 8px;
                }
                th {
                    background-color: #4CAF50;
                    color: white;
                }";
            File.WriteAllText(cssFilePath, customCss);

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Prepare the <link> tag to reference the external stylesheet
            string linkTag = $@"<link rel=""stylesheet"" type=""text/css"" href=""{cssFilePath}"">";

            // Insert the <link> tag just before the closing </head> tag
            int headCloseIndex = htmlContent.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
            if (headCloseIndex >= 0)
            {
                string beforeHeadClose = htmlContent.Substring(0, headCloseIndex);
                string afterHeadClose = htmlContent.Substring(headCloseIndex);
                htmlContent = beforeHeadClose + linkTag + Environment.NewLine + afterHeadClose;
            }
            else
            {
                // If </head> not found, prepend the link at the beginning of the file
                htmlContent = linkTag + Environment.NewLine + htmlContent;
            }

            // Write the modified HTML back to the file
            File.WriteAllText(htmlFilePath, htmlContent);

            Console.WriteLine("HTML export completed with inline styles disabled.");
            Console.WriteLine($"External stylesheet linked: {cssFilePath}");
        }
    }
}
