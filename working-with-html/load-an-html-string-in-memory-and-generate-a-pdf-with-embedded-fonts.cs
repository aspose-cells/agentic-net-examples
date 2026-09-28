// Title: Convert an HTML string to PDF with embedded fonts using Aspose.Cells in C#
// AI Prompts: Generate C# code that reads an HTML snippet from a MemoryStream, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, and saves it as a PDF with fonts embedded using PdfSaveOptions. | Show how to configure Aspose.Cells PdfSaveOptions to guarantee font embedding when converting in‑memory HTML content to a PDF file in a .NET console application. | Write a self‑contained C# console program that converts a UTF‑8 HTML string to a PDF without creating temporary files, using Aspose.Cells HtmlLoadOptions and PdfSaveOptions.
// Common Searches: asp.net convert html string to pdf with embedded fonts using aspose.cells | c# load html from memory stream into workbook and export to pdf | how to guarantee font embedding when saving pdf from html in aspose.cells | aspose.cells pdfsaveoptions font embedding example c#
// Tags: Aspose.Cells HTML import | PdfSaveOptions embedded fonts | in-memory HTML to PDF conversion C# | MemoryStream workbook loading Aspose.Cells | export workbook as PDF with fonts embedded

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// // Loads an HTML string into a MemoryStream, creates a Workbook with HtmlLoadOptions, and saves it as a PDF using PdfSaveOptions (fonts are embedded by default).
class HtmlToPdfConverter
{
    static void Main()
    {
        try
        {
            // HTML content to be converted
            string htmlContent = "<html><body><h1>Hello, Aspose.Cells!</h1><p>This is a sample HTML.</p></body></html>";

            // Convert HTML string to a byte array and load it into a memory stream
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                // Load the HTML into a Workbook using HtmlLoadOptions
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Configure PDF save options (fonts are embedded by default in recent versions)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file
                workbook.Save("output.pdf", pdfOptions);
            }

            Console.WriteLine("PDF file 'output.pdf' created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
