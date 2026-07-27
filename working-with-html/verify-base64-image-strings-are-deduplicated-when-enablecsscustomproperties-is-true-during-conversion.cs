// Title: Check Base64 Image Deduplication with EnableCssCustomProperties in Aspose.Cells HTML Export (C#)
// Description: A C# example that adds the same PNG picture to two cells, saves the workbook as HTML with ExportImagesAsBase64 and EnableCssCustomProperties enabled, then reads the output to verify that the base64 data URI appears only once, confirming CSS‑based deduplication.
// Keywords: Aspose.Cells | HtmlSaveOptions | EnableCssCustomProperties | base64 image deduplication | ExportImagesAsBase64 | C# HTML conversion | duplicate picture handling | CSS custom properties
// Common Searches: Aspose.Cells deduplicate base64 images | EnableCssCustomProperties duplicate picture | count base64 occurrences in Aspose HTML output | C# verify image deduplication Aspose.Cells | HTML export with CSS custom properties Aspose
// Developer Intent: Confirm that identical images are emitted a single time in the HTML when CSS custom properties are used during conversion.
// Use Cases: Generate compact HTML from a workbook that contains repeated pictures. | Programmatically validate that a single base64 data URI is referenced via a CSS variable. | Inspect the generated CSS custom property to ensure correct embedding of the image.
// AI Prompts: Create a C# unit test that adds the same image to multiple cells, saves to HTML with EnableCssCustomProperties true, and asserts a single occurrence of the base64 string. | Write a C# method that extracts the CSS custom property containing the embedded base64 image from Aspose.Cells HTML output. | Explain the mechanism Aspose.Cells uses to deduplicate base64 images when EnableCssCustomProperties is enabled.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// A C# example that adds the same PNG picture to two cells, saves the workbook as HTML with ExportImagesAsBase64 and EnableCssCustomProperties enabled, then reads the output to verify that the base64 data URI appears only once, confirming CSS‑based deduplication.
class HtmlBase64DeduplicationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample PNG image (1x1 red pixel) as base64 string
        const string redPixelBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(redPixelBase64);

        // Add the same image twice to different cells
        using (MemoryStream ms1 = new MemoryStream(imageBytes))
        {
            sheet.Pictures.Add(1, 1, ms1);
        }
        using (MemoryStream ms2 = new MemoryStream(imageBytes))
        {
            sheet.Pictures.Add(5, 3, ms2);
        }

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;               // Embed images as base64
        htmlOptions.EnableCssCustomProperties = true;          // Enable deduplication via CSS custom properties

        // Save the workbook to HTML
        const string outputHtml = "deduplication_demo.html";
        workbook.Save(outputHtml, htmlOptions);

        // Load the generated HTML
        string htmlContent = File.ReadAllText(outputHtml);

        // Count how many times the original base64 string appears in the HTML
        int occurrenceCount = 0;
        int startIndex = 0;
        while ((startIndex = htmlContent.IndexOf(redPixelBase64, startIndex, StringComparison.Ordinal)) != -1)
        {
            occurrenceCount++;
            startIndex += redPixelBase64.Length;
        }

        // Output verification result
        Console.WriteLine($"Base64 image data occurrence count in HTML: {occurrenceCount}");
        if (occurrenceCount == 1)
        {
            Console.WriteLine("Success: Base64 image string is deduplicated when EnableCssCustomProperties is true.");
        }
        else
        {
            Console.WriteLine("Failure: Base64 image string appears multiple times.");
        }

        // Optional: Show a snippet of the generated CSS custom property (for manual inspection)
        var cssMatch = Regex.Match(htmlContent, @"--[^:]+:\s*url\(['""]?data:image/[^;]+;base64,[^'"")]+\)");
        if (cssMatch.Success)
        {
            Console.WriteLine("Detected CSS custom property containing the base64 image:");
            Console.WriteLine(cssMatch.Value);
        }
    }
}
