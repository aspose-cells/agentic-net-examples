// Title: Deduplicate Identical Images in HTML Export with EnableCssCustomProperties – Aspose.Cells for .NET
// Description: This C# example creates a workbook, inserts the same PNG into two cells, and saves it as HTML with ExportImagesAsBase64 and EnableCssCustomProperties turned on. After saving, the code parses the HTML, extracts all data‑image Base64 URIs, and verifies that only one distinct Base64 string is emitted while multiple <img> tags exist, proving CSS‑custom‑property deduplication.
// Keywords: Aspose.Cells HTML export | EnableCssCustomProperties | base64 image deduplication | C# .NET spreadsheet to HTML | duplicate image handling | CSS custom properties Aspose | reduce HTML size Aspose.Cells | global developers | North America .NET | Europe C#
// Common Searches: Aspose.Cells duplicate images HTML export | EnableCssCustomProperties base64 example C# | how to deduplicate images in Aspose HTML output | count distinct base64 strings Aspose.Cells | verify image deduplication Aspose.Cells .NET
// Developer Intent: Ensure that identical pictures are emitted once as a Base64 URI and referenced via CSS custom properties during HTML conversion.
// Use Cases: Automated regression test that confirms image deduplication reduces HTML payload. | Generating compact HTML reports or email templates where a logo appears multiple times. | Validating compliance with size‑budget constraints for web‑published spreadsheet exports.
// AI Prompts: Create an xUnit test in C# that adds the same image twice to a workbook, saves to HTML with EnableCssCustomProperties=true, and asserts that the distinct Base64 count equals 1. | Write a PowerShell script that scans an Aspose.Cells‑generated HTML file, lists all data:image Base64 URIs, and flags duplicates. | Explain the internal mechanism Aspose.Cells uses to replace repeated Base64 images with CSS custom properties during HTML export.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, inserts the same PNG into two cells, and saves it as HTML with ExportImagesAsBase64 and EnableCssCustomProperties turned on. After saving, the code parses the HTML, extracts all data‑image Base64 URIs, and verifies that only one distinct Base64 string is emitted while multiple <img> tags exist, proving CSS‑custom‑property deduplication.
    public class VerifyBase64ImageDeduplication
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the image file
                string imagePath = "logo.png";

                // Ensure the image file exists before adding
                if (File.Exists(imagePath))
                {
                    // Add the same image to two different cells
                    int imgIndex1 = sheet.Pictures.Add(1, 1, imagePath);
                    Picture pic1 = sheet.Pictures[imgIndex1];
                    pic1.Width = 100;
                    pic1.Height = 100;

                    int imgIndex2 = sheet.Pictures.Add(5, 3, imagePath);
                    Picture pic2 = sheet.Pictures[imgIndex2];
                    pic2.Width = 100;
                    pic2.Height = 100;
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping image insertion.");
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true,
                    EnableCssCustomProperties = true
                };

                // Save the workbook as HTML
                string htmlPath = "OutputWithCssCustomProperties.html";
                workbook.Save(htmlPath, htmlOptions);

                // Load the generated HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Find all Base64 image data URIs in the HTML
                Regex base64Regex = new Regex(@"data:image\/[a-zA-Z]+;base64,[A-Za-z0-9+/=]+", RegexOptions.Compiled);
                MatchCollection matches = base64Regex.Matches(htmlContent);

                // Count distinct Base64 strings
                HashSet<string> distinctBase64 = new HashSet<string>();
                foreach (Match match in matches)
                {
                    distinctBase64.Add(match.Value);
                }

                Console.WriteLine($"Total <img> tags with Base64 data: {matches.Count}");
                Console.WriteLine($"Distinct Base64 image strings: {distinctBase64.Count}");

                // Verification: when EnableCssCustomProperties is true, identical images should be stored once
                if (distinctBase64.Count == 1 && matches.Count > 1)
                {
                    Console.WriteLine("Verification passed: Base64 image strings are deduplicated using CSS custom properties.");
                }
                else
                {
                    Console.WriteLine("Verification failed: Image deduplication did not occur as expected.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyBase64ImageDeduplication.Run();
        }
    }
}
