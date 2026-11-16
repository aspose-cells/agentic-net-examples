// Title: Convert an HTML file to PDF while preserving line spacing and paragraph indentation with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert HTML to PDF keep original line spacing | preserve paragraph indentation when saving HTML as PDF using Aspose.Cells | how to maintain HTML layout after PDF export with Aspose.Cells .NET | Aspose.Cells HtmlLoadOptions PDF conversion retain formatting C# | save HTML document as PDF with exact spacing using Aspose.Cells library
// Tags: Aspose.Cells HtmlLoadOptions PDF conversion | Workbook.Save preserving document structure | C# HTML to PDF conversion retaining vertical spacing | Aspose.Cells maintain formatting during PDF export | HTML to PDF export layout fidelity .NET | Aspose.Cells PDF output exact spacing

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions, creates the output folder if needed, and saves the workbook as a PDF with SaveFormat.Pdf, preserving the original line spacing and paragraph indentation while handling any errors.
class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = @"C:\Input\sample.html";

        // Path for the output PDF file
        string pdfPath = @"C:\Output\sample.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Load the HTML document into a workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}
