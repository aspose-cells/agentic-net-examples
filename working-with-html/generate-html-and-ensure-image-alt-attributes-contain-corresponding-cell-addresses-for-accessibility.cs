// Title: Add cell address alt text to images when exporting an Excel worksheet to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, saves the first worksheet as HTML, and inserts an alt attribute containing each picture’s upper‑left cell reference into the corresponding <img> tags. | Create a C# routine that builds a dictionary mapping Aspose.Cells Picture objects to their exported image filenames and uses a regular expression to add or replace alt attributes in the generated HTML for accessibility compliance. | Implement a C# utility that verifies the workbook exists, exports images as external files, reads the resulting HTML, and updates every <img> element with the cell address of the picture as the alt value.
// Common Searches: how to set alt attribute for images exported by Aspose.Cells HTML save in C# | C# Aspose.Cells map picture to cell address for accessible HTML output | regex replace img tags with cell reference alt text after Aspose.Cells conversion | generate accessible HTML from Excel workbook using Aspose.Cells and add alt text to pictures
// Tags: Aspose.Cells HTML export with image alt attributes | C# picture-to-cell mapping for Excel images | regex replace img tags for accessibility | add alt text from picture upper‑left cell | export Excel images as external files using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The program loads an Excel workbook, creates a dictionary linking each picture’s exported image file name (e.g., image1.png) to the cell address of its upper‑left corner, saves the first worksheet as HTML with external image files, reads the generated HTML, uses a regular expression to locate <img> tags, injects an alt attribute containing the corresponding cell address (removing any existing alt), and writes the modified HTML to a new file for improved accessibility.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string htmlPath = "output.html";
            const string accessibleHtmlPath = "output_accessible.html";

            // Ensure the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Verify there is at least one worksheet.
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Work with the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Map image file name -> cell address where the image is placed.
            var imageAltMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int pictureCounter = 0;

            foreach (Picture picture in sheet.Pictures)
            {
                pictureCounter++;

                // Cell address of the picture's upper‑left corner.
                string cellAddress = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);

                // Aspose.Cells names exported images as image1.png, image2.png, …
                string imageFileName = $"image{pictureCounter}.png";

                imageAltMap[imageFileName] = cellAddress;
            }

            // Save the workbook as HTML (images as external files).
            var htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = false,
                ExportActiveWorksheetOnly = true
            };
            workbook.Save(htmlPath, htmlOptions);

            // Verify the HTML file was created.
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Failed to generate HTML file \"{htmlPath}\".");
                return;
            }

            // Load the generated HTML.
            string htmlContent = File.ReadAllText(htmlPath);

            // Regex to locate <img> tags and capture the src attribute.
            string imgPattern = @"<img\s+([^>]*?)src\s*=\s*""([^""]+)""([^>]*?)>";
            htmlContent = Regex.Replace(htmlContent, imgPattern, match =>
            {
                string beforeSrc = match.Groups[1].Value;
                string srcValue = match.Groups[2].Value;
                string afterSrc = match.Groups[3].Value;

                // Extract just the file name from the src (it may contain a path).
                string fileName = Path.GetFileName(srcValue);

                // Determine the appropriate alt text (cell address) if we have a mapping.
                string altAttribute = string.Empty;
                if (imageAltMap.TryGetValue(fileName, out string cellAddr) && !string.IsNullOrEmpty(cellAddr))
                {
                    altAttribute = $" alt=\"{cellAddr}\"";
                }

                // Remove any existing alt attribute from the remaining part of the tag.
                string cleanedAfterSrc = Regex.Replace(afterSrc, @"\s+alt\s*=\s*""[^""]*""", "", RegexOptions.IgnoreCase);

                // Reconstruct the <img> tag with the new alt attribute.
                return $"<img {beforeSrc}src=\"{srcValue}\"{cleanedAfterSrc}{altAttribute}>";
            }, RegexOptions.IgnoreCase);

            // Save the modified HTML with accessible alt attributes.
            File.WriteAllText(accessibleHtmlPath, htmlContent);
            Console.WriteLine($"Accessible HTML saved to \"{accessibleHtmlPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
