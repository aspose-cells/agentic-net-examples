// Title: C# unit test to verify that Aspose.Cells HtmlSaveOptions.DisableCss removes all <style> tags from generated HTML
// Common Searches: how to write a C# unit test for Aspose.Cells HtmlSaveOptions.DisableCss to check for style tags | verify that Aspose.Cells HTML export does not include <style> elements when DisableCss is true | C# Aspose.Cells unit test example for HTML output without CSS style blocks | NUnit test to ensure no style tags appear in Aspose.Cells generated HTML | check generated HTML from Aspose.Cells for absence of CSS when DisableCss is enabled
// Tags: Aspose.Cells HtmlSaveOptions DisableCss verification | C# unit test HTML export without CSS | Aspose.Cells HTML output style tag removal | NUnit Aspose.Cells HTML generation test | validate no CSS style blocks in generated HTML

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates a C# unit test that creates a workbook, saves it to HTML with HtmlSaveOptions.DisableCss set to true, reads the output, and asserts that no <style> tags are present, confirming CSS suppression in the generated HTML.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add some data
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Text");
                sheet.Cells["B2"].PutValue(42);

                // Configure HTML save options to disable CSS generation
                var saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    DisableCss = true
                };

                // Save the workbook to a memory stream as HTML
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, saveOptions);
                    ms.Position = 0;

                    // Read the generated HTML content
                    string html;
                    using (var reader = new StreamReader(ms, Encoding.UTF8))
                    {
                        html = reader.ReadToEnd();
                    }

                    // Verify that no <style> tags are present in the output
                    bool containsStyleTag = html.IndexOf("<style", StringComparison.OrdinalIgnoreCase) >= 0;
                    if (containsStyleTag)
                    {
                        Console.WriteLine("Test Failed: HTML output contains <style> tags.");
                    }
                    else
                    {
                        Console.WriteLine("Test Passed: No <style> tags found.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
        }
    }
}
