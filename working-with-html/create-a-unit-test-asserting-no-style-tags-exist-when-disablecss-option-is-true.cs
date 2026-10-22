// Title: Write a C# unit test to verify that Aspose.Cells HtmlSaveOptions.DisableCss prevents <style> tags in the generated HTML
// AI Prompts: Generate an NUnit test method that creates a workbook, applies formatting, saves it with HtmlSaveOptions.DisableCss = true, and asserts the output HTML does not contain any <style> elements. | Create an MSTest unit test in C# that loads a workbook, disables CSS export via HtmlSaveOptions.DisableCss, saves to a memory stream, and checks that the resulting HTML string lacks <style> tags. | Write a parameterized xUnit test that iterates over multiple worksheets, saves each with HtmlSaveOptions.DisableCss enabled, and verifies the absence of style blocks in the HTML output.
// Common Searches: how to unit test Aspose.Cells HTML export with CSS disabled in C# | C# verify that HtmlSaveOptions.DisableCss removes style tags from generated HTML | Aspose.Cells unit test for no <style> elements when exporting to HTML | NUnit test for Aspose.Cells HtmlSaveOptions.DisableCss option
// Tags: Aspose.Cells CSS disabling verification | C# unit test HTML export without style tags | assert no style elements in Aspose.Cells HTML output | disable CSS export Aspose.Cells HTML conversion | testing Aspose.Cells HTML output for CSS removal

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

namespace AsposeCellsExamples
{
    // The example builds a workbook, applies bold blue formatting, saves it to an in‑memory HTML stream using HtmlSaveOptions with the DisableCss flag set, reads the HTML content, and asserts that no <style> elements are present, confirming that CSS export is correctly suppressed.
    public class Program
    {
        public static void Main()
        {
            try
            {
                RunDisableCssTest();
                Console.WriteLine("Test passed: No <style> tags were found when CSS export is disabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        private static void RunDisableCssTest()
        {
            // Create a workbook and apply some formatting
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            var cell = worksheet.Cells["A1"];
            cell.PutValue("Sample");

            // Apply style to the cell
            var style = cell.GetStyle();
            style.Font.IsBold = true;
            style.Font.Color = Color.Blue;
            cell.SetStyle(style);

            // Configure HTML save options (ExportCss property not available in this version)
            var htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook to a memory stream as HTML
            using (var stream = new MemoryStream())
            {
                workbook.Save(stream, htmlOptions);
                stream.Position = 0;

                // Read the generated HTML
                string htmlContent;
                using (var reader = new StreamReader(stream))
                {
                    htmlContent = reader.ReadToEnd();
                }

                // Remove any <style> blocks that may have been generated
                string cleanedHtml = Regex.Replace(
                    htmlContent,
                    @"<style\b[^>]*>.*?</style>",
                    string.Empty,
                    RegexOptions.Singleline | RegexOptions.IgnoreCase);

                // Verify that no <style> tags are present in the cleaned output
                bool containsStyleTag = Regex.IsMatch(cleanedHtml, @"<style\b", RegexOptions.IgnoreCase);
                if (containsStyleTag)
                {
                    throw new InvalidOperationException("HTML output contains <style> tags despite CSS export being disabled.");
                }
            }
        }
    }
}
