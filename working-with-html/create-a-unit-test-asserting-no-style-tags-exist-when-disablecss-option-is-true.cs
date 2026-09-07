// Title: Create a C# unit test that verifies Aspose.Cells HTML export with DisableCss enabled produces HTML without <style> elements
// AI Prompts: Generate a C# test method (using NUnit, MSTest, or xUnit) that builds a workbook, applies a style, saves it to HTML with HtmlSaveOptions.DisableCss = true, reads the HTML from a MemoryStream, and asserts that the string does not contain any <style> tags. | Write code to programmatically export a styled Aspose.Cells workbook to HTML with CSS disabled, capture the output, and fail the test if a <style> element is detected.
// Common Searches: Aspose.Cells unit test for DisableCss option in C# | How to assert that exported HTML from Aspose.Cells contains no style tags | C# verify Aspose.Cells HtmlSaveOptions.DisableCss removes CSS | Testing Aspose.Cells HTML output without <style> elements | Write MSTest for Aspose.Cells HTML export with CSS disabled
// Tags: Aspose.Cells HtmlSaveOptions.DisableCss verification | C# Aspose.Cells HTML export validation | verify absence of style elements | memory stream HTML capture Aspose.Cells | disable CSS in Aspose.Cells HTML generation

using System;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example creates a workbook, applies a red font style to a cell, saves it to HTML with HtmlSaveOptions.DisableCss set to true, reads the HTML from a memory stream, and throws an exception if any <style> tag is found.
    public class Program
    {
        public static void Main()
        {
            try
            {
                DisableCss_ShouldNotContainStyleTags();
                Console.WriteLine("Test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
            }
        }

        private static void DisableCss_ShouldNotContainStyleTags()
        {
            // Create a new workbook and add some styled content
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            var style = workbook.CreateStyle();
            style.Font.Color = Color.Red;
            sheet.Cells["A1"].SetStyle(style);

            // Configure HTML save options with DisableCss = true
            var htmlOptions = new HtmlSaveOptions
            {
                DisableCss = true,
                ExportImagesAsBase64 = true
            };

            // Save the workbook to a memory stream as HTML
            using (var memoryStream = new MemoryStream())
            {
                workbook.Save(memoryStream, htmlOptions);
                memoryStream.Position = 0;

                // Read the generated HTML content
                string htmlContent;
                using (var reader = new StreamReader(memoryStream, Encoding.UTF8))
                {
                    htmlContent = reader.ReadToEnd();
                }

                // Verify that no <style> tags exist in the output
                if (htmlContent.Contains("<style", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("When DisableCss is true, the exported HTML should not contain any <style> tags.");
                }
            }
        }
    }
}
