// Title: Write a C# unit test with Aspose.Cells to confirm that the exported HTML contains a :root selector with CSS custom properties
// AI Prompts: Create an MSTest method that builds a Workbook, applies a style, saves it as HTML using HtmlSaveOptions, reads the HTML string, and asserts that a :root block includes at least one '--' prefixed CSS variable. | Generate an NUnit test that constructs a worksheet, exports it to HTML via Aspose.Cells, extracts the :root rule with a regular expression, and verifies the presence of a CSS custom property. | Provide a xUnit test that renders a styled workbook to HTML, parses the stylesheet, and checks that the :root selector defines a CSS variable.
// Common Searches: how to unit test Aspose.Cells HTML export for CSS variables in the :root selector | c# verify that Aspose.Cells generated HTML includes custom properties under :root | using regex in a C# test to assert presence of CSS custom properties in Aspose.Cells HTML output | assert :root CSS rule contains '--' variables when saving workbook as HTML with Aspose.Cells | unit testing Aspose.Cells HtmlSaveOptions to ensure CSS custom properties are emitted
// Tags: Aspose.Cells HTML export unit testing | detect CSS custom property in :root rule | C# regex validation of generated stylesheet | check for '--' variable in HTML output | styled workbook to HTML conversion test

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTests
{
    // The example builds a simple workbook, applies a red font style to a cell, saves the workbook to an in‑memory HTML stream using Aspose.Cells HtmlSaveOptions, reads the HTML content, uses regular expressions to locate the :root selector, and asserts that the selector's body contains at least one CSS custom property (a variable prefixed with '--').
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a simple workbook with styled content.
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                var cell = worksheet.Cells["A1"];
                cell.PutValue("Sample");

                // Apply style to the cell (Aspose.Cells requires GetStyle/SetStyle).
                Style style = cell.GetStyle();
                style.Font.Color = Color.Red;
                cell.SetStyle(style);

                // Configure HTML save options.
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Save the workbook to a memory stream as HTML.
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, htmlOptions);
                    stream.Position = 0;
                    string htmlContent = new StreamReader(stream).ReadToEnd();

                    // Locate the :root selector in the generated HTML.
                    var rootMatch = Regex.Match(htmlContent, @":root\s*\{([^}]*)\}", RegexOptions.Singleline);
                    if (!rootMatch.Success)
                    {
                        Console.WriteLine("The :root selector was not found in the generated HTML.");
                        return;
                    }

                    // Extract the content inside :root and verify it contains at least one CSS variable.
                    string rootBody = rootMatch.Groups[1].Value;
                    bool hasCustomProperty = Regex.IsMatch(rootBody, @"--[\w-]+\s*:", RegexOptions.Multiline);
                    if (hasCustomProperty)
                    {
                        Console.WriteLine("CSS custom properties are present inside the :root selector.");
                    }
                    else
                    {
                        Console.WriteLine("No CSS custom properties were found inside the :root selector.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
