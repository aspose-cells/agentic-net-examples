// Title: Add a Custom CSS Class to the <body> After Exporting an Aspose.Cells Workbook to HTML (C#)
// Description: C# example that creates a workbook, saves it as a single HTML file with embedded CssStyles, then injects a custom class into the <body> tag to apply additional styling.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | SaveAsSingleFile | CssStyles | add body class | inject CSS into exported HTML | post‑export HTML manipulation | custom CSS class | HTML workbook export
// Common Searches: how to add a CSS class to body after Aspose.Cells HTML export | Aspose.Cells C# add body class to generated HTML | inject custom CSS into single HTML file saved by Aspose.Cells | modify <body> tag in Aspose.Cells HTML output | post‑save HTML editing Aspose.Cells
// Developer Intent: Insert a custom CSS class into the <body> tag of the HTML file produced by Aspose.Cells to enable additional styling.
// Use Cases: Apply corporate branding by attaching a specific class to the body of exported reports. | Enable responsive layout or theme switching through a body class that CSS frameworks can target. | Allow JavaScript modules to locate the exported document via a known body class. | Combine embedded CssStyles with external stylesheet references for layered styling.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a single HTML file and then adds a custom class attribute to the <body> tag without breaking the existing content. | Show how to use HtmlAgilityPack in C# to programmatically add a CSS class to the <body> element of an Aspose.Cells‑generated HTML file. | Explain how to merge Aspose.Cells CssStyles with a post‑save body‑class injection to support both inline styles and external CSS frameworks. | Provide a PowerShell script that automates adding a body class to multiple HTML files exported from Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // C# example that creates a workbook, saves it as a single HTML file with embedded CssStyles, then injects a custom class into the <body> tag to apply additional styling.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Save as a single HTML file so that CssStyles are applied
                SaveAsSingleFile = true,

                // Define additional CSS that targets a custom class on the body element
                CssStyles = @"
                    body.my-custom-class {
                        font-family: Arial, sans-serif;
                        background-color: #f0f0f0;
                        padding: 10px;
                    }"
            };

            // Path for the intermediate HTML file
            string tempHtmlPath = Path.Combine(Path.GetTempPath(), "temp_output.html");

            // Save the workbook as HTML
            workbook.Save(tempHtmlPath, htmlOptions);

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(tempHtmlPath);

            // Insert the custom CSS class into the <body> tag
            // This simple replace works because the file is saved as a single HTML document
            string updatedHtml = htmlContent.Replace("<body>", "<body class=\"my-custom-class\">");

            // Write the modified HTML back to the same file (or a new file if preferred)
            File.WriteAllText(tempHtmlPath, updatedHtml);

            Console.WriteLine("HTML file with custom body class created at:");
            Console.WriteLine(tempHtmlPath);
        }
    }
}
