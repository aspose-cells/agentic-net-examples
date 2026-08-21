// Title: C# – Export Aspose.Cells Workbook to HTML and Add a Custom CSS Class to Header Cells
// Description: Creates a workbook, marks the first row as headers, saves it to HTML with ExportRowColumnHeadings, injects a <style> block for a custom‑header class, replaces every <th> with <th class="custom-header">, and writes the modified file.
// Keywords: Aspose.Cells HTML export | C# add CSS class to th | ExportRowColumnHeadings | custom header styling | modify generated HTML | Aspose.Cells HtmlSaveOptions | inject CSS into Aspose.Cells output
// Common Searches: how to add a CSS class to table headers after Aspose.Cells HTML export | Aspose.Cells C# inject custom style into <th> elements | add custom-header class to Aspose.Cells generated HTML | modify Aspose.Cells HTML output with custom CSS | export workbook to HTML with styled header cells
// Developer Intent: Apply a custom CSS class to all <th> elements in the HTML produced by Aspose.Cells, enabling consistent header styling.
// Use Cases: Brand HTML reports with company colors by styling table headers. | Improve readability of spreadsheet‑derived tables on web pages. | Enable JavaScript libraries to target header cells for sorting or filtering.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML, then adds a <style> block defining a custom header class and injects that class into every <th> tag. | Show how to use HtmlSaveOptions.ExportRowColumnHeadings to export header rows as <th> before applying custom CSS. | Provide a method to read the generated HTML file, insert a CSS rule for a custom header, replace <th> tags with a class attribute, and save the updated file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlHeaderClassDemo
{
    // Creates a workbook, marks the first row as headers, saves it to HTML with ExportRowColumnHeadings, injects a <style> block for a custom‑header class, replaces every <th> with <th class="custom-header">, and writes the modified file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row (row 0) and some data rows
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // Configure HTML save options to export row/column headings
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportRowColumnHeadings = true   // ensure header cells are exported as <th>
            };

            // First save the workbook to HTML
            string htmlPath = "output.html";
            workbook.Save(htmlPath, saveOptions);

            // Load the generated HTML, add a custom CSS class to all header cells (<th>)
            string htmlContent = File.ReadAllText(htmlPath);
            // Insert a custom CSS rule for the new class
            string customCss = @"
                <style>
                    th.custom-header { background-color:#f0f0f0; font-weight:bold; }
                </style>";
            // Place the custom CSS just before the closing </head> tag (or after <head>)
            if (htmlContent.Contains("</head>"))
                htmlContent = htmlContent.Replace("</head>", customCss + "\n</head>");
            else
                htmlContent = customCss + "\n" + htmlContent;

            // Add the class attribute to every <th> element
            htmlContent = htmlContent.Replace("<th>", "<th class=\"custom-header\">");

            // Save the modified HTML to a new file
            string modifiedHtmlPath = "output_with_custom_header.html";
            File.WriteAllText(modifiedHtmlPath, htmlContent);

            Console.WriteLine("HTML saved to: " + htmlPath);
            Console.WriteLine("Modified HTML with custom header class saved to: " + modifiedHtmlPath);
        }
    }
}
