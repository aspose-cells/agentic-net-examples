// Title: C# – Export Aspose.Cells Workbook to HTML with CSS Only (No Inline Styles) and Verify
// Description: Creates an Excel workbook, applies bold red formatting, saves it as HTML using Aspose.Cells HtmlSaveOptions with CSS enabled (DisableCss = false), reads the generated file, counts "style=\"" occurrences, and confirms that no inline style attributes are present.
// Keywords: Aspose.Cells HTML export C# | disable inline styles Aspose.Cells | HtmlSaveOptions DisableCss false | verify HTML output Aspose.Cells | count style attributes C# | Excel to HTML CSS only | clean HTML from Excel | Aspose.Cells CSS separation
// Common Searches: Aspose.Cells export HTML without inline style attributes C# | How to disable inline styles when saving Excel as HTML with Aspose.Cells | Count style="" occurrences in generated HTML C# | HtmlSaveOptions CSS only Aspose.Cells example | Validate Aspose.Cells HTML output for inline styles
// Developer Intent: Generate HTML from an Excel workbook using Aspose.Cells with external CSS only and programmatically confirm that no inline style attributes exist.
// Use Cases: Produce SEO‑friendly, maintainable HTML reports from Excel data by separating styling into CSS. | Add an automated test to CI pipelines that ensures Aspose.Cells HTML exports contain zero inline styles. | Create web‑ready pages from styled worksheets while complying with accessibility and style‑separation best practices.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as HTML using external CSS only and verifies that the output contains zero "style=\"" attributes. | Provide a method to parse a generated HTML file and return the count of inline style attributes, based on Aspose.Cells HtmlSaveOptions settings. | Explain how to configure HtmlSaveOptions to embed CSS in a separate file and programmatically confirm that no inline styles remain in the HTML.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsHtmlInlineCheck
{
    // Creates an Excel workbook, applies bold red formatting, saves it as HTML using Aspose.Cells HtmlSaveOptions with CSS enabled (DisableCss = false), reads the generated file, counts "style=\"" occurrences, and confirms that no inline style attributes are present.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add some formatted data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Bold Red Text");
                Style style = sheet.Cells["A1"].GetStyle();
                style.Font.IsBold = true;
                style.Font.Color = Color.Red;
                sheet.Cells["A1"].SetStyle(style);

                // Configure HTML save options to use CSS (disable inline styles)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    DisableCss = false // CSS enabled, inline styles should not be generated
                };
                string htmlFile = "output.html";

                // Save the workbook as HTML
                workbook.Save(htmlFile, saveOptions);

                // Verify that the HTML file was created
                if (!File.Exists(htmlFile))
                {
                    Console.WriteLine($"Error: HTML file '{htmlFile}' was not created.");
                    return;
                }

                // Load the generated HTML content
                string htmlContent = File.ReadAllText(htmlFile);

                // Count occurrences of inline style attributes
                int inlineStyleCount = 0;
                int index = 0;
                while ((index = htmlContent.IndexOf("style=\"", index, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    inlineStyleCount++;
                    index += 7; // Move past the found occurrence
                }

                // Output the result
                Console.WriteLine($"Number of inline style attributes found: {inlineStyleCount}");
                if (inlineStyleCount == 0)
                {
                    Console.WriteLine("Success: No inline style attributes exist when CSS is enabled.");
                }
                else
                {
                    Console.WriteLine("Failure: Inline style attributes were found despite CSS being enabled.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
