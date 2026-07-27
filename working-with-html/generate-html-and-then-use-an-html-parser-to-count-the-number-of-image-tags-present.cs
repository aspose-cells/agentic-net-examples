// Title: C# Aspose.Cells: Export Workbook to HTML and Count <img> Tags
// Description: Creates a new Workbook, inserts up to two picture files, saves the sheet as an HTML file with external image files, reads the generated markup, and uses a case‑insensitive regular expression to count the <img> elements. Includes file‑existence checks and error handling for robust execution.
// Keywords: Aspose.Cells HTML export C# | count img tags .NET | Excel to HTML with images | Aspose.Cells picture export | regex image tag count C# | C# workbook to HTML | Aspose.Cells HtmlSaveOptions | programmatic image counting
// Common Searches: how to export Excel to HTML with images using Aspose.Cells | C# count <img> tags in generated HTML file | Aspose.Cells save images as separate files | regex to find img tags in .NET | sample code for Aspose.Cells HTML conversion
// Developer Intent: Generate an HTML representation of an Excel workbook and determine how many <img> elements were produced.
// Use Cases: Verify that every picture added to a worksheet appears in the HTML output by matching the <img> count to the number of inserted images. | Include image‑tag verification in automated regression suites for Excel‑to‑HTML conversion pipelines. | Produce a quick inventory of images in exported HTML files for documentation, SEO audits, or content analysis.
// AI Prompts: Provide a C# example that saves an Aspose.Cells workbook as HTML with external image files and then counts the <img> tags using a regular expression. | Suggest a more reliable technique than regex for counting <img> elements in the generated HTML, handling self‑closing tags and attribute variations. | Explain how to configure HtmlSaveOptions to embed images as Base64 and discuss the impact on counting <img> tags in that scenario.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHtmlImageCount
{
    // Creates a new Workbook, inserts up to two picture files, saves the sheet as an HTML file with external image files, reads the generated markup, and uses a case‑insensitive regular expression to count the <img> elements. Includes file‑existence checks and error handling for robust execution.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Paths to sample images
                string imagePath1 = "image1.png";
                string imagePath2 = "image2.png";

                // Add images only if the files exist to avoid FileNotFoundException
                if (File.Exists(imagePath1))
                {
                    worksheet.Pictures.Add(0, 0, imagePath1);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file not found: {imagePath1}");
                }

                if (File.Exists(imagePath2))
                {
                    worksheet.Pictures.Add(5, 5, imagePath2);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file not found: {imagePath2}");
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export images as separate files so that <img> tags are generated
                    ExportImagesAsBase64 = false
                };

                // Define output HTML file path
                string htmlPath = "output.html";

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, htmlOptions);

                // Ensure the HTML file was created
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine("Error: HTML file was not generated.");
                    return;
                }

                // Load the generated HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Count the number of <img> tags using a simple regex (case‑insensitive)
                int imgTagCount = Regex.Matches(htmlContent, "<img\\b", RegexOptions.IgnoreCase).Count;

                // Output the result
                Console.WriteLine($"Number of <img> tags in the generated HTML: {imgTagCount}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
