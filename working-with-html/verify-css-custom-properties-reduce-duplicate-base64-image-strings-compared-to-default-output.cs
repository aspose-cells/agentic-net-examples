// Title: Aspose.Cells C# – Compare Base64 Image Duplication With and Without CSS Custom Properties
// Description: A C# demo that inserts the same PNG into two cells, exports the workbook to HTML twice (EnableCssCustomProperties false/true), extracts data‑URI strings with regex, and reports total vs. unique Base64 images to show whether CSS custom properties eliminate duplicate image data.
// Keywords: Aspose.Cells HTML export | C# base64 image duplication | EnableCssCustomProperties | CSS custom properties Aspose | reduce duplicate data URI | HTMLSaveOptions base64 images | image embedding optimization
// Common Searches: Aspose.Cells CSS custom properties duplicate base64 images | How to check image data‑URI duplication in Aspose.Cells HTML output | EnableCssCustomProperties effect on HTML size | C# count unique base64 images Aspose.Cells | Remove repeated base64 strings with CSS variables
// Developer Intent: Verify if enabling CSS custom properties during HTML export consolidates identical images into a single variable and reduces duplicate Base64 strings.
// Use Cases: Validate the size benefit of CSS custom properties for repeated images in exported HTML. | Generate a quick report of total versus unique Base64 image URIs for quality checks. | Automate comparison of two export configurations to choose the most compact HTML output.
// AI Prompts: Create a C# function that reads an Aspose.Cells‑generated HTML file and returns the count of total and distinct Base64 image data‑URIs. | Write a PowerShell script that runs the Aspose.Cells export with EnableCssCustomProperties true and false, then logs the file size difference and duplicate count. | Explain how Aspose.Cells leverages CSS custom properties to reference repeated images and the impact on HTML payload.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCssCustomPropertiesDemo
{
    // A C# demo that inserts the same PNG into two cells, exports the workbook to HTML twice (EnableCssCustomProperties false/true), extracts data‑URI strings with regex, and reports total vs. unique Base64 images to show whether CSS custom properties eliminate duplicate image data.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a small PNG image (1x1 red pixel) as a byte array
            // This base64 string represents a valid PNG image
            string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lKXcVwAAAABJRU5ErkJggg==";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            using (MemoryStream imgStream = new MemoryStream(imageBytes))
            {
                // Add the same image to two different cells
                // First occurrence at cell B2 (row 1, column 1)
                sheet.Pictures.Add(1, 1, imgStream);
                // Reset stream position for the second addition
                imgStream.Position = 0;
                // Second occurrence at cell E5 (row 4, column 4)
                sheet.Pictures.Add(4, 4, imgStream);
            }

            // Ensure images are exported as Base64 strings
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.ExportImagesAsBase64 = true;

            // Save HTML with CSS custom properties disabled (default)
            options.EnableCssCustomProperties = false;
            string pathWithoutCss = "HtmlWithoutCssCustomProperties.html";
            workbook.Save(pathWithoutCss, options);

            // Save HTML with CSS custom properties enabled
            options.EnableCssCustomProperties = true;
            string pathWithCss = "HtmlWithCssCustomProperties.html";
            workbook.Save(pathWithCss, options);

            // Load the generated HTML files
            string htmlWithoutCss = File.ReadAllText(pathWithoutCss);
            string htmlWithCss = File.ReadAllText(pathWithCss);

            // Regex to capture data URI of embedded images
            Regex dataUriRegex = new Regex(@"data:image\/[^;]+;base64,[A-Za-z0-9+/=]+", RegexOptions.Compiled);

            // Analyze the file without CSS custom properties
            var matchesWithout = dataUriRegex.Matches(htmlWithoutCss);
            int totalOccurrencesWithout = matchesWithout.Count;
            var uniqueWithout = new System.Collections.Generic.HashSet<string>();
            foreach (Match m in matchesWithout) uniqueWithout.Add(m.Value);
            int uniqueCountWithout = uniqueWithout.Count;

            // Analyze the file with CSS custom properties
            var matchesWith = dataUriRegex.Matches(htmlWithCss);
            int totalOccurrencesWith = matchesWith.Count;
            var uniqueWith = new System.Collections.Generic.HashSet<string>();
            foreach (Match m in matchesWith) uniqueWith.Add(m.Value);
            int uniqueCountWith = uniqueWith.Count;

            // Output the comparison results
            Console.WriteLine("=== Comparison of Base64 Image Embedding ===");
            Console.WriteLine($"File without CSS custom properties: {pathWithoutCss}");
            Console.WriteLine($"  Total image data URIs: {totalOccurrencesWithout}");
            Console.WriteLine($"  Unique image data URIs: {uniqueCountWithout}");
            Console.WriteLine();
            Console.WriteLine($"File with CSS custom properties: {pathWithCss}");
            Console.WriteLine($"  Total image data URIs: {totalOccurrencesWith}");
            Console.WriteLine($"  Unique image data URIs: {uniqueCountWith}");
            Console.WriteLine();
            Console.WriteLine("Observation:");
            if (uniqueCountWith < totalOccurrencesWith)
                Console.WriteLine("- CSS custom properties reduced duplicate Base64 strings.");
            else
                Console.WriteLine("- No reduction observed; duplicate strings may still exist.");
        }
    }
}
