// Title: How to convert an HTML file to PDF and embed title and author metadata using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an HTML file with Aspose.Cells, assigns BuiltInDocumentProperties for Title and Author, and saves it as a PDF. | Write a C# snippet that reads metadata values (Title, Author, Subject) from a JSON configuration file, applies them to a Workbook loaded from HTML, and exports the result to PDF using Aspose.Cells. | Show how to programmatically create the output directory if it does not exist before saving the PDF generated from HTML with Aspose.Cells.
// Common Searches: Aspose.Cells C# convert HTML to PDF with custom PDF metadata | add custom PDF metadata using Aspose.Cells HtmlLoadOptions | C# example for adding PDF metadata after converting HTML with Aspose.Cells | how to ensure output folder exists before saving PDF with Aspose.Cells in .NET
// Tags: Aspose.Cells HTML to PDF with built‑in document properties | C# set PDF title author using Aspose.Cells | HtmlLoadOptions workbook conversion to PDF | save workbook as PDF with custom metadata | ensure output directory exists C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the presence of an input HTML file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, sets the built‑in document properties Title and Author, creates the output folder if needed, and saves the workbook as a PDF while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string htmlPath = "input.html";
            string pdfPath = "output.pdf";

            // Verify that the HTML source file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: Input file not found – {htmlPath}");
                return;
            }

            // Load the HTML file into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Embed document metadata using built‑in properties
            workbook.BuiltInDocumentProperties.Title = "Sample Document Title";
            workbook.BuiltInDocumentProperties.Author = "John Doe";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to: {pdfPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
