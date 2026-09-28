// Title: Convert HTML containing <video> elements to PDF with static image placeholders using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML file, uses a regular expression to replace every <video> tag with a specified PNG image, and saves the result as a PDF via Aspose.Cells. | Show how to load the modified HTML content from a MemoryStream into an Aspose.Cells Workbook and export it to PDF with custom image dimensions. | Add robust error handling for missing HTML or placeholder image files when performing the HTML‑to‑PDF conversion with Aspose.Cells.
// Common Searches: how to replace <video> tags with images before converting HTML to PDF using Aspose.Cells in C# | Aspose.Cells load HTML from MemoryStream and save as PDF example | C# regex to substitute video elements with placeholder PNG for PDF export
// Tags: Aspose.Cells HTML to PDF conversion | video tag substitution with PNG in C# | MemoryStream HTML loading Aspose.Cells | SaveFormat.Pdf workbook export | static image placeholder for video elements

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program reads an HTML file, replaces each <video> element with a PNG placeholder via regex, loads the modified HTML into an Aspose.Cells Workbook using a MemoryStream, and saves the workbook as a PDF.
class HtmlToPdfWithVideoPlaceholders
{
    static void Main()
    {
        // Paths for input HTML, placeholder image, and output PDF.
        const string htmlPath = "input.html";
        const string placeholderImagePath = "video_placeholder.png";
        const string pdfPath = "output.pdf";

        try
        {
            // Verify required files exist.
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"HTML file not found: {htmlPath}");
            if (!File.Exists(placeholderImagePath))
                throw new FileNotFoundException($"Placeholder image not found: {placeholderImagePath}");

            // Read the HTML content.
            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // Replace each <video> element with an <img> placeholder.
            // The placeholder image is referenced by its file name; adjust the path if needed.
            string pattern = @"<video[\s\S]*?</video>";
            string replacement = $"<img src=\"{placeholderImagePath}\" width=\"320\" height=\"240\" alt=\"Video placeholder\" />";
            string modifiedHtml = Regex.Replace(htmlContent, pattern, replacement, RegexOptions.IgnoreCase);

            // Load the modified HTML into a workbook using a memory stream.
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(modifiedHtml)))
            {
                var loadOptions = new LoadOptions(LoadFormat.Html);
                var workbook = new Workbook(stream, loadOptions);

                // Save the workbook as PDF.
                workbook.Save(pdfPath, SaveFormat.Pdf);
            }

            Console.WriteLine($"PDF generated successfully: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
