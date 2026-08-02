// Title: Export Aspose.Cells Workbook to HTML and Verify CSS Custom Properties in a :root Selector
// Description: C# sample that creates an Aspose.Cells workbook, applies bold red styling, optionally embeds images, and saves it as a single HTML file with HtmlSaveOptions.EnableCssCustomProperties enabled. The generated HTML is read from a memory stream, inspected for a ":root" selector to confirm CSS custom properties, and written to disk.
// Keywords: Aspose.Cells HTML export | EnableCssCustomProperties | C# generate HTML from Excel | CSS custom properties root selector | single HTML file Aspose.Cells | embed images Aspose.Cells HTML | verify CSS variables in output | Aspose.Cells HtmlSaveOptions
// Common Searches: how to enable CSS custom properties when saving Excel to HTML with Aspose.Cells | check for :root selector in Aspose.Cells generated HTML | save workbook as single HTML file using Aspose.Cells .NET | Aspose.Cells HTMLSaveOptions embed CSS variables | C# verify CSS variables in exported HTML
// Developer Intent: Export an Excel workbook to HTML and ensure the output uses CSS custom properties defined in a :root selector.
// Use Cases: Create a compact, theme‑able HTML representation of a styled worksheet for web integration. | Programmatically validate that Aspose.Cells emitted CSS variables by scanning the HTML for a :root block. | Include worksheet images directly in the HTML while keeping styling centralized via CSS custom properties.
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook as HTML with EnableCssCustomProperties turned on and checks for a :root selector. | Write a method to parse the HTML output from Aspose.Cells and extract all CSS custom property definitions from the :root rule. | Explain how to configure HtmlSaveOptions to embed worksheet images as base64 data URIs while preserving CSS custom properties.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlCssCustomPropertiesDemo
{
    // C# sample that creates an Aspose.Cells workbook, applies bold red styling, optionally embeds images, and saves it as a single HTML file with HtmlSaveOptions.EnableCssCustomProperties enabled. The generated HTML is read from a memory stream, inspected for a ":root" selector to confirm CSS custom properties, and written to disk.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Text");

                // Apply bold red font to A1
                var style = sheet.Cells["A1"].GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = System.Drawing.Color.Red;
                sheet.Cells["A1"].SetStyle(style);

                // Add an image if the file exists
                string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.png");
                if (File.Exists(imagePath))
                {
                    int imgIdx1 = sheet.Pictures.Add(1, 1, imagePath);
                    sheet.Pictures[imgIdx1].Width = 100;
                    sheet.Pictures[imgIdx1].Height = 100;

                    int imgIdx2 = sheet.Pictures.Add(5, 3, imagePath);
                    sheet.Pictures[imgIdx2].Width = 100;
                    sheet.Pictures[imgIdx2].Height = 100;
                }

                // Configure HTML save options to enable CSS custom properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    EnableCssCustomProperties = true, // Optimize using CSS custom properties
                    SaveAsSingleFile = true           // Embed CSS for easier parsing
                };

                // Save the workbook to a memory stream as HTML
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, htmlOptions);
                    htmlStream.Position = 0;

                    // Read the generated HTML content
                    string htmlContent;
                    using (StreamReader reader = new StreamReader(htmlStream))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Simple check for a :root selector indicating CSS custom properties
                    bool rootSelectorFound = htmlContent.Contains(":root");

                    Console.WriteLine(rootSelectorFound
                        ? "CSS custom properties are defined in a :root selector."
                        : "No :root selector with CSS custom properties found.");

                    // Write the HTML to a file for inspection
                    string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputWithCssCustomProperties.html");
                    File.WriteAllText(outputPath, htmlContent);
                    Console.WriteLine($"HTML file saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
