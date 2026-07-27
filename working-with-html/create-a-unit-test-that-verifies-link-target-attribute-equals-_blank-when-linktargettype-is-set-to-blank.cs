// Title: C# unit test for Aspose.Cells HtmlLinkTargetType.Blank verifies target="_blank" in exported HTML
// Description: Shows how to build a Workbook, add a hyperlink, set HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank, export to HTML, and assert that the generated markup contains the target="_blank" attribute. Includes exception handling and cleanup, and works with NUnit, MSTest, or xUnit.
// Keywords: Aspose.Cells | HtmlSaveOptions | LinkTargetType | Blank | target=_blank | C# unit test | hyperlink export | HTML output | NUnit | MSTest | xUnit | continuous integration | regression test
// Common Searches: Aspose.Cells unit test HtmlLinkTargetType | verify target blank in HTML export | C# test for hyperlink target attribute | HtmlSaveOptions LinkTargetType.Blank example | how to assert _blank in Aspose.Cells HTML | unit testing Aspose.Cells HTML output
// Developer Intent: Write an automated test that confirms Aspose.Cells adds target="_blank" when HtmlSaveOptions.LinkTargetType is set to Blank.
// Use Cases: Validate that external links open in a new tab after Excel‑to‑HTML conversion | Integrate link‑target verification into CI pipelines to catch regressions | Provide documentation examples for developers using Aspose.Cells HTML export options | Ensure compliance with web accessibility guidelines that require explicit link targets
// AI Prompts: Generate an NUnit test that creates a Workbook, adds a hyperlink, sets HtmlSaveOptions.LinkTargetType to Blank, saves to a MemoryStream, and asserts the HTML contains target="_blank". | Write a MSTest method that verifies Aspose.Cells HtmlSaveOptions with LinkTargetType.Blank produces the correct target attribute without writing to disk. | Provide an xUnit test snippet that checks the exported HTML from Aspose.Cells for target='_blank' when using HtmlLinkTargetType.Blank, including cleanup of temporary resources. | Create a parameterized C# unit test that runs the HtmlLinkTargetType verification for multiple target types (Blank, Self, Parent) and asserts the expected markup.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Shows how to build a Workbook, add a hyperlink, set HtmlSaveOptions.LinkTargetType = HtmlLinkTargetType.Blank, export to HTML, and assert that the generated markup contains the target="_blank" attribute. Includes exception handling and cleanup, and works with NUnit, MSTest, or xUnit.
    public class HtmlLinkTargetTypeDemo
    {
        public static void Main()
        {
            // Arrange: create a workbook with a hyperlink
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Visit Example");
            worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

            // Set HTML save options to use the Blank target type
            var saveOptions = new HtmlSaveOptions
            {
                LinkTargetType = HtmlLinkTargetType.Blank
            };

            // Define a temporary HTML file path
            string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");

            try
            {
                // Act: save the workbook as HTML
                workbook.Save(tempHtmlPath, saveOptions);

                // Verify the generated HTML contains target="_blank"
                if (File.Exists(tempHtmlPath))
                {
                    string htmlContent = File.ReadAllText(tempHtmlPath);
                    if (htmlContent.Contains("target=\"_blank\""))
                    {
                        Console.WriteLine("Success: target=\"_blank\" found in generated HTML.");
                    }
                    else
                    {
                        Console.WriteLine("Failure: target=\"_blank\" not found in generated HTML.");
                    }
                }
                else
                {
                    Console.WriteLine("Error: HTML file was not created.");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: log any unexpected exceptions
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
                        // Ignore cleanup errors
                    }
                }
            }
        }
    }
}
