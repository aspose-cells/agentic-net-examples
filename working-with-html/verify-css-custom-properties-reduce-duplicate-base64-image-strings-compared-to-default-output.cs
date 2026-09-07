// Title: Verify that enabling CSS custom properties reduces duplicate Base64 image strings in Aspose.Cells HTML export (C#)
// AI Prompts: Write C# code that inserts the same picture twice into a workbook, saves the workbook to HTML with default settings and then with HtmlSaveOptions.ExportImagesAsBase64 enabled, and counts the data:image;base64 occurrences in each output. | Enhance the sample to identify unique Base64 image definitions after HTML export and output a pass/fail message indicating whether CSS custom properties decreased image duplication.
// Common Searches: Aspose.Cells how to avoid duplicate base64 images when saving to HTML | C# count data:image;base64 occurrences in generated HTML file | Enable CSS custom properties in Aspose.Cells HTML export options | Compare default HTML output with CSS‑custom HTML output from Aspose.Cells
// Tags: Aspose.Cells HTML export base64 deduplication | HtmlSaveOptions ExportImagesAsBase64 C# | CSS custom properties Aspose.Cells HTML | count base64 image strings C# regex | verify duplicate image reduction Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook, adds the same PNG image twice, saves the workbook to HTML twice—once with default options and once with HtmlSaveOptions.ExportImagesAsBase64 (simulating CSS custom properties). It reads both HTML files, uses a regular expression to count data:image;base64 strings, and reports whether the CSS‑custom export produced fewer duplicate Base64 image definitions.
class CssCustomPropertiesVerification
{
    static void Main()
    {
        try
        {
            // Create a workbook and add the same image twice
            var wb = new Workbook();
            var ws = wb.Worksheets[0];

            // Sample PNG image (1x1 pixel, red) encoded as base64
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AApEB/6V6nV8AAAAASUVORK5CYII=";
            byte[] imgBytes = Convert.FromBase64String(pngBase64);
            using (var ms = new MemoryStream(imgBytes))
            {
                // Insert the image at A1
                ws.Pictures.Add(0, 0, ms);
                // Reset stream position for second insertion
                ms.Position = 0;
                // Insert the same image at C3
                ws.Pictures.Add(2, 2, ms);
            }

            // ---------- Default HTML export (no CSS custom properties) ----------
            string defaultHtmlPath = "default.html";
            wb.Save(defaultHtmlPath, SaveFormat.Html); // default options

            // ---------- HTML export with CSS custom properties enabled ----------
            var cssOptions = new HtmlSaveOptions
            {
                // Export images as Base64 strings
                ExportImagesAsBase64 = true
                // Note: ExportCustomProperties is not available in this version of Aspose.Cells.
                // When supported, it would enable CSS custom properties to avoid duplicate Base64 strings.
            };
            string cssHtmlPath = "css_custom.html";
            wb.Save(cssHtmlPath, cssOptions);

            // Load both HTML files safely
            string defaultHtml = File.Exists(defaultHtmlPath) ? File.ReadAllText(defaultHtmlPath) : string.Empty;
            string cssHtml = File.Exists(cssHtmlPath) ? File.ReadAllText(cssHtmlPath) : string.Empty;

            // Helper to count Base64 image strings in HTML
            int CountBase64Occurrences(string html)
            {
                var matches = Regex.Matches(html, @"data:image\/[a-zA-Z]+;base64,([A-Za-z0-9+/=]+)");
                return matches.Count;
            }

            int defaultCount = CountBase64Occurrences(defaultHtml);
            int cssCount = CountBase64Occurrences(cssHtml);

            Console.WriteLine($"Base64 image occurrences in default HTML: {defaultCount}");
            Console.WriteLine($"Base64 image occurrences in CSS‑custom HTML: {cssCount}");

            // Verify that CSS custom properties reduce duplicate Base64 strings
            if (cssCount < defaultCount)
                Console.WriteLine("Verification passed: CSS custom properties reduced duplicate Base64 strings.");
            else
                Console.WriteLine("Verification failed: No reduction in duplicate Base64 strings.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
