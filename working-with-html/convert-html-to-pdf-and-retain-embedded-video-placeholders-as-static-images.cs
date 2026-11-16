// Title: Convert HTML to PDF in C# with Aspose.Cells while substituting video elements with a static placeholder image
// Common Searches: C# Aspose.Cells replace <video> tags with an image before exporting HTML to PDF | How to convert HTML containing videos to PDF using Aspose.Cells and show placeholder images | Aspose.Cells HTML to PDF conversion ignoring video content by inserting a PNG placeholder | Replace video elements with <img> tags in HTML for PDF generation using Aspose.Cells .NET | Load modified HTML from a temporary file into an Aspose.Cells workbook and save as PDF
// Tags: Aspose.Cells HTML to PDF with custom image substitution | C# regex replace video tags with img placeholder | load modified HTML into Aspose.Cells workbook | save workbook as PDF using SaveFormat.Pdf | temporary HTML file handling for Aspose.Cells conversion

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML file, uses a regular expression to replace all <video> elements with an <img> tag that points to a specified placeholder PNG, writes the updated HTML to a temporary file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, saves the workbook as a PDF, and then cleans up the temporary file.
class HtmlToPdfConverter
{
    static void Main(string[] args)
    {
        try
        {
            // Input HTML file path
            string htmlPath = @"C:\Input\sample.html";

            // Output PDF file path
            string pdfPath = @"C:\Output\sample.pdf";

            // Placeholder image to represent video content
            string placeholderImagePath = @"C:\Resources\video_placeholder.png";

            // Verify required files exist
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file not found: {htmlPath}");
                return;
            }

            if (!File.Exists(placeholderImagePath))
            {
                Console.WriteLine($"Placeholder image not found: {placeholderImagePath}");
                return;
            }

            // Read the original HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Replace <video>...</video> blocks with an <img> tag pointing to the placeholder image
            string videoPattern = @"<video\b[^>]*>.*?<\/video>";
            string replacement = $"<img src=\"{placeholderImagePath}\" alt=\"Video placeholder\" />";
            string modifiedHtml = Regex.Replace(htmlContent, videoPattern, replacement,
                                                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Write modified HTML to a temporary file (Aspose.Cells loads from file)
            string tempHtmlPath = Path.Combine(Path.GetTempPath(),
                                               $"temp_{Guid.NewGuid()}.html");
            File.WriteAllText(tempHtmlPath, modifiedHtml);

            // Load HTML into a workbook
            var loadOptions = new HtmlLoadOptions();
            var workbook = new Workbook(tempHtmlPath, loadOptions);

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Clean up temporary file
            if (File.Exists(tempHtmlPath))
                File.Delete(tempHtmlPath);

            Console.WriteLine($"PDF successfully created at: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
