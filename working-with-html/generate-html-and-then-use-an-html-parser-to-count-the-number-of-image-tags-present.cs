// Title: C# – Export Aspose.Cells Workbook to HTML with Base64 Images and Count <img> Tags
// Description: Creates a workbook, inserts a picture, saves it as HTML with images embedded as Base64 via HtmlSaveOptions, reads the generated file, and counts the <img> elements using a case‑insensitive regular expression (or an HTML parser).
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportImagesAsBase64 | embed Excel picture | count img tags | regex image count | HtmlAgilityPack | Excel to self‑contained HTML | automated HTML validation
// Common Searches: Aspose.Cells export workbook to HTML base64 | C# count <img> elements in generated HTML | how to embed Excel picture as Base64 using Aspose | HtmlSaveOptions ExportImagesAsBase64 sample code | parse HTML and count images with C#
// Developer Intent: Generate a self‑contained HTML report from an Excel workbook and determine how many image tags appear in the output.
// Use Cases: Verify that pictures added to a worksheet are correctly embedded in the HTML export. | Create email‑ready HTML reports that do not rely on external image files. | Automate regression tests for Excel‑to‑HTML conversion by comparing expected and actual <img> counts. | Build a lightweight utility that audits Excel workbooks for missing or extra images after conversion.
// AI Prompts: Write C# code that adds a picture to an Aspose.Cells worksheet, saves the workbook as HTML with ExportImagesAsBase64 enabled, and counts the <img> tags using a case‑insensitive regex. | Explain the impact of ExportImagesAsBase64 on the HTML output and suggest a robust method (e.g., HtmlAgilityPack) to count image elements without false positives. | Provide a C# snippet that parses the saved HTML with HtmlAgilityPack and returns the total number of <img> nodes.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlImageCount
{
    // Creates a workbook, inserts a picture, saves it as HTML with images embedded as Base64 via HtmlSaveOptions, reads the generated file, and counts the <img> elements using a case‑insensitive regular expression (or an HTML parser).
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the sample image. Ensure the file exists before adding.
                string imagePath = "example.jpg";
                if (File.Exists(imagePath))
                {
                    // Add the image at cell A1 (row 0, column 0).
                    worksheet.Pictures.Add(0, 0, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping image insertion.");
                }

                // Configure HTML save options to embed images as Base64.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as an HTML file.
                string htmlPath = "output.html";
                workbook.Save(htmlPath, htmlOptions);

                // Verify the HTML file was created.
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Failed to create HTML file: {htmlPath}");
                    return;
                }

                // Read the generated HTML content.
                string htmlContent = File.ReadAllText(htmlPath);

                // Count <img> tags using a simple regex (case‑insensitive).
                int imageTagCount = Regex.Matches(htmlContent, "<img\\b", RegexOptions.IgnoreCase).Count;

                // Output the result.
                Console.WriteLine($"Number of <img> tags in the HTML: {imageTagCount}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
