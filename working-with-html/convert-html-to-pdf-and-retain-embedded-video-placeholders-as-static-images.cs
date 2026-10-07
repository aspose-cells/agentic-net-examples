// Title: Convert HTML containing <video> elements to PDF with static image placeholders using Aspose.Cells for .NET
// AI Prompts: Read an HTML file, replace every <video> tag with an <img> tag that points to a placeholder PNG, and save the result as a PDF using Aspose.Cells in C#. | Create a C# console application that checks for the presence of the source HTML and placeholder image files, substitutes video elements with <img> tags via a regular expression, and generates a PDF through Aspose.Cells. | Add comprehensive error handling to the conversion workflow so that missing files are reported and any unexpected exceptions are logged while producing a PDF with video placeholders rendered as static images.
// Common Searches: c# replace video tags with image before converting html to pdf using aspose.cells | how to use Aspose.Cells to convert html file to pdf while showing video placeholders as pictures | regex pattern to remove <video> elements from html for pdf generation in .net | Aspose.Cells html to pdf conversion preserving layout with static video thumbnails
// Tags: video element substitution with placeholder graphic Aspose.Cells | pdf generation from processed html Aspose.Cells | pattern driven video tag removal C# | load processed html into Aspose.Cells workbook | embedded video thumbnail rendering in pdf

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an input HTML file, uses a regular expression to replace each <video> element with an <img> tag that references a specified placeholder PNG, writes the modified HTML to a temporary file, loads it into an Aspose.Cells Workbook, and saves the workbook as a PDF. It includes file‑existence checks and exception handling to ensure robust conversion while rendering video sections as static images.
class HtmlToPdfWithVideoPlaceholders
{
    static void Main()
    {
        try
        {
            // Paths for input HTML, placeholder image, and output PDF
            string inputHtmlPath = "input.html";
            string placeholderImagePath = "video_placeholder.png";
            string outputPdfPath = "output.pdf";

            // Verify required files exist
            if (!File.Exists(inputHtmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {inputHtmlPath}");
            if (!File.Exists(placeholderImagePath))
                throw new FileNotFoundException($"Placeholder image file not found: {placeholderImagePath}");

            // Read the original HTML content
            string htmlContent = File.ReadAllText(inputHtmlPath);

            // Replace <video>...</video> tags with an <img> tag pointing to the placeholder image
            string videoPattern = @"<video\b[^>]*>.*?</video>";
            string imgReplacement = $"<img src=\"{placeholderImagePath}\" alt=\"Video placeholder\" />";
            string processedHtml = Regex.Replace(
                htmlContent,
                videoPattern,
                imgReplacement,
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Write the processed HTML to a temporary file (useful for debugging)
            string tempHtmlPath = "processed.html";
            File.WriteAllText(tempHtmlPath, processedHtml);

            // Load the processed HTML into an Aspose.Cells Workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(tempHtmlPath, loadOptions);

            // Save the workbook as PDF; placeholder images appear where videos were
            workbook.Save(outputPdfPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
