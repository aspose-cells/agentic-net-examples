// Title: C# Unit Test: Verify Aspose.Cells HTML Export Produces No <style> Tags When DisableCss Is Enabled
// Description: Creates a workbook, applies bold formatting, configures HtmlSaveOptions with DisableCss = true, saves to a MemoryStream as HTML, reads the output, and asserts that the generated HTML contains no <style> elements. Ideal for automated testing of inline‑style only HTML export in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | HtmlSaveOptions.DisableCss | C# unit test | HTML export without CSS | verify no style tags | inline styles only | Aspose.Cells HTML output test | MSTest | NUnit | XUnit
// Common Searches: Aspose.Cells unit test for DisableCss | How to check that exported HTML has no <style> tags in C# | HtmlSaveOptions.DisableCss example | Testing Aspose.Cells HTML output for inline styles | C# verify Aspose.Cells HTML does not contain CSS blocks
// Developer Intent: Write an automated test that confirms Aspose.Cells generates HTML without any <style> elements when the DisableCss option is turned on.
// Use Cases: Validate HTML for email templates that prohibit embedded style blocks. | Ensure compliance with strict Content Security Policy (CSP) rules that block <style> tags. | Add a regression test to detect future changes in Aspose.Cells HTML rendering behavior.
// AI Prompts: Generate an MSTest method that creates a workbook, sets HtmlSaveOptions.DisableCss = true, saves to a string, and asserts the string does not contain '<style>'. | Provide an XUnit test snippet that logs the HTML output from Aspose.Cells and fails if any '<style>' tag is found. | Write a mock‑free NUnit test for Aspose.Cells HTML export that checks for the absence of CSS blocks in CI pipelines.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, applies bold formatting, configures HtmlSaveOptions with DisableCss = true, saves to a MemoryStream as HTML, reads the output, and asserts that the generated HTML contains no <style> elements. Ideal for automated testing of inline‑style only HTML export in Aspose.Cells for .NET.
    public class HtmlDisableCssDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add styled content
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                var cell = worksheet.Cells["A1"];
                cell.PutValue("Styled Text");
                var style = cell.GetStyle();
                style.Font.IsBold = true;
                cell.SetStyle(style);

                // Configure HtmlSaveOptions to disable CSS (use only inline styles)
                var saveOptions = new HtmlSaveOptions
                {
                    DisableCss = true
                };

                // Save the workbook to a memory stream as HTML
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, saveOptions);
                    stream.Position = 0;
                    string htmlContent = new StreamReader(stream).ReadToEnd();

                    // Verify that the generated HTML does not contain any <style> tags
                    bool containsStyleTag = htmlContent.Contains("<style", StringComparison.OrdinalIgnoreCase);
                    Console.WriteLine(containsStyleTag
                        ? "Test Failed: HTML contains <style> tags."
                        : "Test Passed: HTML does not contain <style> tags.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
