// Title: Unit Test: Verify HtmlLinkTargetType.Parent Generates target="_parent" in Aspose.Cells HTML Export
// Description: Creates a workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent, saves as HTML to a memory stream, and asserts that the generated anchor tag contains the target="_parent" attribute.
// Keywords: Aspose.Cells | HtmlSaveOptions | LinkTargetType | Parent | _parent | hyperlink target attribute | C# unit test | HTML export validation | MSTest | xUnit | NUnit
// Common Searches: Aspose.Cells unit test for hyperlink target | HtmlLinkTargetType Parent example | verify target=_parent in generated HTML | C# test HtmlSaveOptions LinkTargetType | Aspose.Cells HTML export hyperlink target
// Developer Intent: Write an automated test that confirms the HTML output from Aspose.Cells includes target="_parent" when HtmlSaveOptions.LinkTargetType is set to Parent.
// Use Cases: Continuous‑integration regression test for HTML export behavior | Ensuring frame navigation works correctly after library updates | Validating custom hyperlink target settings in multi‑frame web applications
// AI Prompts: Generate an MSTest method that creates a workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent, saves to a MemoryStream, and asserts the HTML contains target="_parent". | Provide an xUnit test for Aspose.Cells that verifies the anchor tag includes target="_parent" when using HtmlLinkTargetType.Parent. | Write a NUnit test example that checks Aspose.Cells HTML output for the correct target attribute based on the selected LinkTargetType.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent, saves as HTML to a memory stream, and asserts that the generated anchor tag contains the target="_parent" attribute.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Add some text and a hyperlink to cell A1
                worksheet.Cells["A1"].PutValue("Aspose");
                worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

                // Set HTML save options to use Parent target type (target="_parent")
                var saveOptions = new HtmlSaveOptions
                {
                    LinkTargetType = HtmlLinkTargetType.Parent
                };

                // Save the workbook to a memory stream as HTML
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, saveOptions);
                    stream.Position = 0;

                    // Read the generated HTML content
                    string htmlContent;
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Verify that the anchor tag contains target="_parent"
                    if (htmlContent.Contains("target=\"_parent\""))
                    {
                        Console.WriteLine("Success: target=\"_parent\" attribute found in generated HTML.");
                    }
                    else
                    {
                        Console.WriteLine("Failure: target=\"_parent\" attribute not found in generated HTML.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
