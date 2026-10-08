// Title: C# unit test to verify Aspose.Cells HTML export omits <style> tags when CSS class generation is disabled
// AI Prompts: Create a C# test method that saves a Workbook to HTML with HtmlSaveOptions.ExportCssClass set to false, reads the result from a MemoryStream, and asserts that the output string does not contain any <style> elements. | Generate an MSTest/NUnit compatible unit test that uses Aspose.Cells to export a worksheet to HTML, disables CSS class generation via reflection if necessary, and fails the test when a <style> tag is detected in the generated HTML.
// Common Searches: Aspose.Cells HtmlSaveOptions ExportCssClass false unit test C# | how to assert no style tags in HTML output from Aspose.Cells | C# verify that HTML export from Aspose.Cells does not include CSS when disabled | unit testing Aspose.Cells HTML export without embedded style elements | disable CSS class generation in Aspose.Cells HTML export and check output
// Tags: Aspose.Cells HtmlSaveOptions disable CSS class generation | C# unit test for HTML export without style tags | validate Aspose.Cells HTML output using memory stream | assert no <style> elements in generated HTML Aspose.Cells | export workbook to HTML with CSS disabled Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example builds a workbook, applies bold formatting, and saves it to HTML using HtmlSaveOptions with the ExportCssClass property forced to false (via reflection for version compatibility). The HTML is captured from a MemoryStream and the test throws an exception if any <style> tags are present, demonstrating how to unit‑test the DisableCss behavior of Aspose.Cells HTML export.
    public class HtmlExportTests
    {
        public static void Main()
        {
            try
            {
                ExportToHtml_DisableCss_ShouldNotContainStyleTags();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        public static void ExportToHtml_DisableCss_ShouldNotContainStyleTags()
        {
            // Create a new workbook and add some data with formatting
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            var cell = sheet.Cells["A1"];
            cell.PutValue("Header");

            // Apply bold formatting using GetStyle/SetStyle (compatible with all versions)
            var style = cell.GetStyle();
            style.Font.IsBold = true;
            cell.SetStyle(style);

            // Prepare HTML save options with CSS class generation disabled (using reflection for compatibility)
            var htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true
            };

            // Attempt to set ExportCssClass property if it exists in the current Aspose.Cells version
            var exportCssProp = typeof(HtmlSaveOptions).GetProperty("ExportCssClass");
            if (exportCssProp != null && exportCssProp.CanWrite)
            {
                exportCssProp.SetValue(htmlOptions, false);
            }

            // Save the workbook to a memory stream as HTML
            string htmlContent;
            using (var ms = new MemoryStream())
            {
                try
                {
                    workbook.Save(ms, htmlOptions);
                }
                catch (Exception saveEx)
                {
                    throw new InvalidOperationException("Failed to save workbook as HTML.", saveEx);
                }

                ms.Position = 0;
                using (var reader = new StreamReader(ms))
                {
                    htmlContent = reader.ReadToEnd();
                }
            }

            // Verify that the generated HTML does not contain any <style> tags
            if (htmlContent.Contains("<style", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("HTML output should not contain any <style> tags when CSS is disabled.");
            }
        }
    }
}
