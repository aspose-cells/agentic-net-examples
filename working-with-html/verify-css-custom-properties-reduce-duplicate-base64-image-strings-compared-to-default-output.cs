// Title: Aspose.Cells C# – Reduce duplicate Base64 images using CSS custom properties in HTML export
// Description: A C# demo that inserts the same PNG into two cells, saves the workbook to HTML with and without the EnableCssCustomProperties flag, parses the output for data:image Base64 URIs, and shows that CSS custom properties collapse duplicate image strings.
// Keywords: Aspose.Cells | C# HTML export | EnableCssCustomProperties | Base64 image deduplication | CSS custom properties | embedded images | Excel to HTML | image count verification
// Common Searches: Aspose.Cells reduce duplicate Base64 images | EnableCssCustomProperties effect on HTML export | C# count embedded images in generated HTML | Aspose.Cells CSS custom properties example | how to deduplicate images when saving Excel as HTML
// Developer Intent: Verify that turning on CSS custom properties during HTML conversion lowers the number of repeated Base64 image strings.
// Use Cases: Run an automated check that a workbook with identical pictures produces fewer data‑URI occurrences when EnableCssCustomProperties is true. | Create lightweight HTML reports by consolidating repeated images into CSS variables. | Validate Aspose.Cells version upgrades for improved image deduplication in HTML output.
// AI Prompts: Generate a C# unit test that asserts countWithCss < countWithoutCss for duplicate pictures using Aspose.Cells. | Explain how EnableCssCustomProperties transforms repeated Base64 images into CSS variables and show the resulting CSS snippet. | Write a PowerShell script that reads two HTML files, extracts unique data:image strings, and prints the counts.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsCssCustomPropertiesDemo
{
    // A C# demo that inserts the same PNG into two cells, saves the workbook to HTML with and without the EnableCssCustomProperties flag, parses the output for data:image Base64 URIs, and shows that CSS custom properties collapse duplicate image strings.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a small PNG image as a byte array (red dot 1x1)
            // This avoids external file dependencies
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAE/wH+6cVYVQAAAABJRU5ErkJggg==");

            // Add the same image to two different cells (A1 and D5)
            sheet.Pictures.Add(0, 0, new MemoryStream(pngBytes));
            sheet.Pictures.Add(4, 3, new MemoryStream(pngBytes));

            // Configure HTML save options to embed images as Base64
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportImagesAsBase64 = true;

            // Save with CSS custom properties enabled (optimised)
            htmlOptions.EnableCssCustomProperties = true;
            string pathWithCss = "Html_WithCssCustomProperties.html";
            workbook.Save(pathWithCss, htmlOptions);

            // Save with CSS custom properties disabled (default behaviour)
            htmlOptions.EnableCssCustomProperties = false;
            string pathWithoutCss = "Html_WithoutCssCustomProperties.html";
            workbook.Save(pathWithoutCss, htmlOptions);

            // Read the generated HTML files
            string htmlWithCss = File.ReadAllText(pathWithCss);
            string htmlWithoutCss = File.ReadAllText(pathWithoutCss);

            // Count occurrences of Base64 image data URIs (pattern: data:image/...;base64,)
            int countWithCss = Regex.Matches(htmlWithCss, @"data:image\/[^;]+;base64,").Count;
            int countWithoutCss = Regex.Matches(htmlWithoutCss, @"data:image\/[^;]+;base64,").Count;

            // Output the results
            Console.WriteLine($"Base64 image occurrences with CSS custom properties: {countWithCss}");
            Console.WriteLine($"Base64 image occurrences without CSS custom properties: {countWithoutCss}");

            // Simple verification
            if (countWithCss < countWithoutCss)
                Console.WriteLine("CSS custom properties reduced duplicate Base64 image strings.");
            else
                Console.WriteLine("No reduction observed; check the input or Aspose.Cells version.");

            // Keep console window open when run outside IDE
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
