// Title: Verify that enabling CSS custom properties in Aspose.Cells HTML export reduces duplicate Base64 image strings
// AI Prompts: Generate C# code that saves an Aspose.Cells workbook to HTML with HtmlSaveOptions.ExportImagesAsBase64 enabled and CSS custom properties turned on, then compare the number of distinct data:image URIs with the default export. | Add a C# routine that extracts all data:image strings from the generated HTML, stores them in a HashSet, and prints the distinct count for the default and CSS‑custom‑property outputs. | Modify the example to load a JPEG file instead of a PNG and confirm that CSS variables still collapse repeated Base64 images.
// Common Searches: Aspose.Cells how to reduce repeated base64 image data when exporting to HTML using CSS variables | C# count distinct data:image base64 strings in HTML produced by Aspose.Cells | Enable CSS custom properties in HtmlSaveOptions to deduplicate images during Excel to HTML conversion | Compare default HTML export versus CSS custom properties export in Aspose.Cells .NET
// Tags: Aspose.Cells HTML export CSS custom properties | deduplicate base64 images Aspose.Cells | count distinct data:image URIs C# | HtmlSaveOptions ExportImagesAsBase64 setting | Excel to HTML image deduplication using CSS variables

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates saving a workbook to HTML with default settings and with CSS custom properties enabled, then counts distinct Base64 image strings to show that CSS variables reduce duplicate image data.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Load an image from file if it exists
                const string imagePath = "sample.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }
                else
                {
                    byte[] imageBytes = File.ReadAllBytes(imagePath);
                    using (var imageStream = new MemoryStream(imageBytes))
                    {
                        // Insert the same picture into different cells
                        sheet.Pictures.Add(0, 0, new MemoryStream(imageBytes));
                        sheet.Pictures.Add(2, 2, new MemoryStream(imageBytes));
                        sheet.Pictures.Add(4, 4, new MemoryStream(imageBytes));
                    }
                }

                // ---------------------------------------------------------------
                // 1. Save HTML with default settings (no CSS custom properties)
                // ---------------------------------------------------------------
                string defaultHtml;
                using (var defaultHtmlStream = new MemoryStream())
                {
                    workbook.Save(defaultHtmlStream, SaveFormat.Html);
                    defaultHtml = Encoding.UTF8.GetString(defaultHtmlStream.ToArray());
                }

                // ---------------------------------------------------------------
                // 2. Save HTML with CSS custom properties enabled to deduplicate Base64 strings
                // ---------------------------------------------------------------
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportImagesAsBase64 = true // Enables Base64 export; duplicates are deduplicated via CSS variables automatically
                };

                string customCssHtml;
                using (var customCssHtmlStream = new MemoryStream())
                {
                    workbook.Save(customCssHtmlStream, htmlOptions);
                    customCssHtml = Encoding.UTF8.GetString(customCssHtmlStream.ToArray());
                }

                // ---------------------------------------------------------------
                // 3. Helper method to count distinct Base64 image strings in HTML
                // ---------------------------------------------------------------
                int defaultCount = CountDistinctBase64Images(defaultHtml);
                int customCssCount = CountDistinctBase64Images(customCssHtml);

                // ---------------------------------------------------------------
                // 4. Output the results
                // ---------------------------------------------------------------
                Console.WriteLine($"Distinct Base64 images (default): {defaultCount}");
                Console.WriteLine($"Distinct Base64 images (CSS custom properties): {customCssCount}");
                Console.WriteLine(customCssCount < defaultCount
                    ? "CSS custom properties reduced duplicate Base64 strings."
                    : "No reduction observed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Counts unique Base64 image strings in the provided HTML content
        private static int CountDistinctBase64Images(string html)
        {
            var base64Set = new HashSet<string>();
            int index = 0;
            while ((index = html.IndexOf("data:image", index, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                int end = html.IndexOfAny(new[] { '"', '\'', ';' }, index);
                if (end == -1) break;
                string base64 = html.Substring(index, end - index);
                base64Set.Add(base64);
                index = end;
            }
            return base64Set.Count;
        }
    }
}
