// Title: Add a Custom CSS Class to the <body> Tag After Converting an Aspose.Cells Workbook to HTML (C#)
// Description: This example creates a workbook, saves it as a single HTML file with a custom CSS definition, then reads the generated file and injects or appends a "custom-body" class to the <body> element, handling existing class attributes safely.
// Keywords: Aspose.Cells | C# | HTML conversion | custom CSS class | body tag | post‑process HTML | HtmlSaveOptions | single file export | inject class attribute | modify generated HTML
// Common Searches: Aspose.Cells add class to body after HTML export | C# append CSS class to <body> in generated HTML | how to modify Aspose.Cells HTML output | inject custom CSS into Aspose.Cells HTML file | add custom-body class to HTML produced by Aspose.Cells
// Developer Intent: Inject or append a custom CSS class to the <body> element of HTML produced by Aspose.Cells.
// Use Cases: Standardize styling of Excel‑to‑HTML reports across an application. | Apply a theme or dark‑mode class without altering the original workbook. | Combine with existing body classes for responsive or mobile layouts. | Enable client‑side scripts that target a specific body class after export.
// AI Prompts: Generate a C# method that adds a specified CSS class to the <body> tag of an Aspose.Cells HTML file, preserving any existing classes. | Write a utility that reads an Aspose.Cells HTML output, inserts multiple CSS classes into the body element, and ensures no duplicate entries. | Create a reusable function to post‑process Aspose.Cells HTML, apply a custom stylesheet defined in HtmlSaveOptions, and return the updated HTML content.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlCustomBodyClass
{
    // This example creates a workbook, saves it as a single HTML file with a custom CSS definition, then reads the generated file and injects or appends a "custom-body" class to the <body> element, handling existing class attributes safely.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save as a single HTML file so that CssStyles can be applied
            htmlOptions.SaveAsSingleFile = true;

            // Define additional CSS styles (optional, can be empty)
            // Here we define a CSS class that can be used for the body element
            htmlOptions.CssStyles = @"
                .custom-body {
                    font-family: Arial, sans-serif;
                    background-color: #f0f0f0;
                    padding: 10px;
                }";

            // Save the workbook to an HTML file
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions);

            // After conversion, read the generated HTML and append the custom CSS class to the <body> tag
            string htmlContent = File.ReadAllText(htmlPath);

            // Replace the opening <body> tag with one that includes the custom class
            // Handles possible whitespace after <body> (e.g., <body>, <body >, <body id="...">)
            // Simple approach: find the first occurrence of "<body" and insert the class attribute
            int bodyTagStart = htmlContent.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            if (bodyTagStart >= 0)
            {
                int bodyTagEnd = htmlContent.IndexOf('>', bodyTagStart);
                if (bodyTagEnd > bodyTagStart)
                {
                    // Check if a class attribute already exists
                    string bodyTag = htmlContent.Substring(bodyTagStart, bodyTagEnd - bodyTagStart + 1);
                    if (bodyTag.IndexOf("class=", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Append the custom class to the existing class attribute
                        htmlContent = htmlContent.Replace(
                            "class=\"",
                            "class=\"custom-body ");
                    }
                    else
                    {
                        // Insert a new class attribute before the closing '>'
                        string newBodyTag = bodyTag.Insert(bodyTag.Length - 1, " class=\"custom-body\"");
                        htmlContent = htmlContent.Remove(bodyTagStart, bodyTag.Length)
                                                 .Insert(bodyTagStart, newBodyTag);
                    }
                }
            }

            // Write the modified HTML back to the file
            File.WriteAllText(htmlPath, htmlContent);

            Console.WriteLine("HTML file generated with custom CSS class applied to the <body> element.");
        }
    }
}
