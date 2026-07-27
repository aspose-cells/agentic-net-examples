// Title: C# Unit Test: Verify HtmlLinkTargetType.Parent Generates target="_parent" in Aspose.Cells HTML Export
// Description: Creates a Workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent, saves to a temporary HTML file, reads the output, and asserts that the <a> tag contains target="_parent". The test reports success or failure and removes the temporary file.
// Keywords: Aspose.Cells | HtmlLinkTargetType | Parent target | C# unit test | HTML export verification | hyperlink target attribute | HtmlSaveOptions test | Aspose.Cells HTML output | automated regression test
// Common Searches: Aspose.Cells unit test for link target parent | verify target=_parent in exported HTML Aspose.Cells | C# test HtmlSaveOptions LinkTargetType Parent | how to assert hyperlink target attribute in Aspose.Cells HTML | Aspose.Cells HTML export link target verification
// Developer Intent: Write an automated test that confirms the HTML produced by Aspose.Cells uses target="_parent" when HtmlLinkTargetType is set to Parent.
// Use Cases: Ensure HTML reports from Excel workbooks open links in the parent frame as required by UI design. | Validate compliance with navigation policies by testing Aspose.Cells HTML export settings. | Detect regressions in hyperlink rendering across Aspose.Cells library updates.
// AI Prompts: Generate an MSTest method that creates a Workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent, saves to a temporary HTML file, and asserts the presence of target="_parent". | Provide an xUnit test for Aspose.Cells that verifies HtmlLinkTargetType.Parent produces the correct <a> tag attribute in the saved HTML. | Show code to clean up temporary HTML files after running a unit test for Aspose.Cells HtmlLinkTargetType Parent.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a Workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Parent, saves to a temporary HTML file, reads the output, and asserts that the <a> tag contains target="_parent". The test reports success or failure and removes the temporary file.
    public class HtmlLinkTargetTypeDemo
    {
        public static void Main()
        {
            // Define a temporary HTML file path
            string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");

            try
            {
                // Create a workbook and add a hyperlink
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Aspose");
                worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

                // Set HTML save options to use Parent target (target="_parent")
                var saveOptions = new HtmlSaveOptions
                {
                    LinkTargetType = HtmlLinkTargetType.Parent
                };

                // Save the workbook as HTML
                workbook.Save(tempHtmlPath, saveOptions);

                // Read the generated HTML content
                string htmlContent = File.ReadAllText(tempHtmlPath);

                // Verify that the hyperlink contains target="_parent"
                var match = Regex.Match(
                    htmlContent,
                    @"<a\s+[^>]*target\s*=\s*[""']_parent[""'][^>]*>",
                    RegexOptions.IgnoreCase);

                if (match.Success)
                {
                    Console.WriteLine("Test passed: hyperlink contains target=\"_parent\".");
                }
                else
                {
                    Console.WriteLine("Test failed: hyperlink does not contain target=\"_parent\".");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary file
                if (File.Exists(tempHtmlPath))
                {
                    try
                    {
                        File.Delete(tempHtmlPath);
                    }
                    catch
                    {
                        // Ignore any errors during cleanup
                    }
                }
            }
        }
    }
}
