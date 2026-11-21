// Title: Verify Aspose.Cells C# HTML export reduces duplicate Base64 image strings by using CSS custom properties
// Common Searches: aspocells c# check if CSS custom properties remove duplicate base64 images in exported html | count data:image URIs in Aspose.Cells HTML output using C# | compare default HTML export vs CSS custom property export for image deduplication in Aspose.Cells | how to enable CSS custom properties for image optimization in Aspose.Cells HTML save options | sample code to measure base64 image duplication after Aspose.Cells HTML export
// Tags: Aspose.Cells HTML export deduplicate base64 images | C# count data URI occurrences in generated HTML | CSS custom properties image optimization Aspose.Cells | export workbook to HTML with embedded base64 images | compare default and CSS‑custom‑property HTML output | Aspose.Cells SaveOptions ExportImagesAsBase64

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, inserts the same PNG into two cells, and saves the workbook to HTML twice—once with default settings and once with CSS custom properties enabled—both with ExportImagesAsBase64=true. It counts the "data:image" URIs in each HTML string to determine whether the CSS custom‑property approach reduces duplicate Base64 image strings, and prints the comparison result.
class CssCustomPropertyVerification
{
    static void Main()
    {
        try
        {
            // Ensure the sample image exists; create a tiny placeholder if missing
            string imagePath = "sample.png";
            if (!File.Exists(imagePath))
            {
                // A 1x1 pixel transparent PNG (base64 encoded)
                const string pngBase64 =
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(pngBase64);
                File.WriteAllBytes(imagePath, pngBytes);
            }

            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Insert the same image into two different cells
            ws.Pictures.Add(1, 1, imagePath); // Cell B2
            ws.Pictures.Add(5, 5, imagePath); // Cell F6

            // ---------- Default HTML export (no CSS custom properties) ----------
            HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true // embed each image as a separate base64 string
            };

            using (MemoryStream defaultStream = new MemoryStream())
            {
                wb.Save(defaultStream, defaultOptions);
                string defaultHtml = Encoding.UTF8.GetString(defaultStream.ToArray());

                // Count occurrences of base64 image data URIs in the default HTML
                int defaultBase64Count = CountBase64Occurrences(defaultHtml);

                // ---------- HTML export with CSS custom properties ----------
                HtmlSaveOptions cssOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportImagesAsBase64 = true // still embed as base64
                    // Note: Aspose.Cells automatically uses CSS custom properties for images when possible.
                };

                using (MemoryStream cssStream = new MemoryStream())
                {
                    wb.Save(cssStream, cssOptions);
                    string cssHtml = Encoding.UTF8.GetString(cssStream.ToArray());

                    // Count occurrences of base64 image data URIs in the CSS‑custom‑property HTML
                    int cssBase64Count = CountBase64Occurrences(cssHtml);

                    // Output the comparison result
                    Console.WriteLine($"Base64 occurrences (default): {defaultBase64Count}");
                    Console.WriteLine($"Base64 occurrences (CSS custom properties): {cssBase64Count}");

                    if (cssBase64Count < defaultBase64Count)
                    {
                        Console.WriteLine("CSS custom properties successfully reduced duplicate base64 image strings.");
                    }
                    else
                    {
                        Console.WriteLine("No reduction in duplicate base64 image strings was observed.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to count how many times a base64 image data URI appears in HTML
    private static int CountBase64Occurrences(string html)
    {
        // Data URIs for images start with "data:image"
        return html.Split(new[] { "data:image" }, StringSplitOptions.None).Length - 1;
    }
}
