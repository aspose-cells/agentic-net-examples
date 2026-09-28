// Title: Check that identical pictures are deduplicated into a single base64 string when saving a workbook to HTML with EnableCssCustomProperties in Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds the same image twice to an Aspose.Cells worksheet, saves the workbook as HTML with EnableCssCustomProperties set to true, and programmatically verifies that only one unique base64 image string appears in the HTML output. | Write a C# routine that parses the HTML produced by Aspose.Cells, extracts all data:image;base64 fragments using a regular expression, and determines whether duplicate images are referenced by a single base64 definition.
// Common Searches: Aspose.Cells enable CSS custom properties for HTML export and deduplicate images | C# verify that duplicate pictures are merged into one base64 string in HTML output | extract and count base64 image data from Aspose.Cells generated HTML file | EnableCssCustomProperties true image deduplication Aspose.Cells example
// Tags: htmlsaveoptions enablecsscustomproperties aspocells | deduplicate base64 images aspocells | c# regex extract base64 image data | verify image deduplication html export

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Cells;

// The program inserts the same PNG image twice into an Aspose.Cells workbook, saves the workbook to HTML with EnableCssCustomProperties enabled, extracts all data:image;base64 strings from the resulting HTML using a regex, and confirms that only one distinct base64 string exists, demonstrating image deduplication.
class Base64ImageDeduplicationCheck
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Load an image into a byte array (ensure the file exists)
            const string imagePath = "sample.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: Image file \"{imagePath}\" not found.");
                return;
            }

            byte[] imageData = File.ReadAllBytes(imagePath);

            // Insert the same image twice into the worksheet using a MemoryStream
            using (MemoryStream imgStream1 = new MemoryStream(imageData))
            {
                int pictureIndex1 = sheet.Pictures.Add(0, 0, imgStream1);
                // pictureIndex1 can be used later if needed
            }

            using (MemoryStream imgStream2 = new MemoryStream(imageData))
            {
                int pictureIndex2 = sheet.Pictures.Add(1, 1, imgStream2);
                // pictureIndex2 can be used later if needed
            }

            // Configure HTML save options to enable CSS custom properties (triggers deduplication)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                EnableCssCustomProperties = true
            };

            // Save the workbook to an HTML string using a MemoryStream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, saveOptions);
                htmlStream.Position = 0;
                string htmlContent = new StreamReader(htmlStream).ReadToEnd();

                // Extract all base64 image strings via regex
                Regex base64Regex = new Regex(@"data:image\/[a-zA-Z]+;base64,[A-Za-z0-9+/=]+");
                MatchCollection matches = base64Regex.Matches(htmlContent);

                // Count distinct base64 strings
                HashSet<string> distinctBase64 = new HashSet<string>();
                foreach (Match match in matches)
                {
                    distinctBase64.Add(match.Value);
                }

                // Output verification results
                Console.WriteLine($"Total base64 image occurrences found: {matches.Count}");
                Console.WriteLine($"Distinct base64 image strings: {distinctBase64.Count}");

                if (distinctBase64.Count == 1 && matches.Count > 1)
                {
                    Console.WriteLine("Deduplication successful: the same base64 image is referenced multiple times but defined only once.");
                }
                else
                {
                    Console.WriteLine("Deduplication failed or not applicable.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
