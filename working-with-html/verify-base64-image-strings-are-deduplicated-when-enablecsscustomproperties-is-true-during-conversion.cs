// Title: Check that identical worksheet pictures are deduplicated into a single base64 data URI when EnableCssCustomProperties is true in Aspose.Cells HTML export (C#)
// AI Prompts: Generate C# code that adds the same PNG image twice to an Aspose.Cells worksheet, activates CSS custom properties for HTML conversion, saves the workbook, and checks that only one unique base64 data URI is present. | Update the example to set HtmlSaveOptions.EnableCssCustomProperties = true, then print the number of distinct base64 image strings found in the exported HTML. | Create an NUnit test in C# that asserts the HTML output from Aspose.Cells with CSS custom properties enabled contains a single distinct data:image base64 string for duplicated pictures.
// Common Searches: aspnet aspose.cells html export duplicate images deduplication | css custom properties true base64 image deduplication asp.net | verify single data uri for repeated pictures in Aspose.Cells HTML output | c# extract and compare base64 image strings from Aspose.Cells generated HTML
// Tags: Aspose.Cells HTMLSaveOptions EnableCssCustomProperties | base64 image deduplication Aspose.Cells | duplicate picture handling HTML export | C# extract data URI from generated HTML | unit testing Aspose.Cells image deduplication

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;
using System.Collections.Generic;

// The sample creates a workbook, inserts the same 1x1 PNG image twice at different cells, saves the workbook to an in‑memory HTML string using HtmlSaveOptions, extracts all data:image base64 strings via regex, collects distinct values in a HashSet, and prints success if only one unique base64 image is present, confirming that duplicate pictures are deduplicated when CSS custom properties are enabled.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Small 1x1 PNG (transparent) encoded in base64
            const string pngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6ZkAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(pngBase64);

            // Insert the same image twice into the worksheet
            using (MemoryStream imgStream1 = new MemoryStream(imageBytes))
            {
                // First picture at A1 (row 0, column 0)
                sheet.Pictures.Add(0, 0, imgStream1);
            }

            using (MemoryStream imgStream2 = new MemoryStream(imageBytes))
            {
                // Second picture at C3 (row 2, column 2)
                sheet.Pictures.Add(2, 2, imgStream2);
            }

            // Configure HTML save options (default options are sufficient)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Save the workbook to an in‑memory HTML string
            string htmlContent;
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, saveOptions);
                htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());
            }

            // Extract all base64 image strings from the generated HTML
            // Pattern matches data:image/...;base64,XXXXX
            Regex base64Regex = new Regex(@"data:image\/[a-zA-Z]+;base64,[A-Za-z0-9+/=]+");
            MatchCollection matches = base64Regex.Matches(htmlContent);

            // Collect distinct base64 strings
            HashSet<string> distinctBase64 = new HashSet<string>();
            foreach (Match match in matches)
            {
                distinctBase64.Add(match.Value);
            }

            // Verify deduplication: there should be only one distinct base64 image
            if (distinctBase64.Count == 1)
            {
                Console.WriteLine("Success: Base64 image strings are deduplicated.");
            }
            else
            {
                Console.WriteLine($"Failure: Expected 1 distinct base64 image, found {distinctBase64.Count}.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
