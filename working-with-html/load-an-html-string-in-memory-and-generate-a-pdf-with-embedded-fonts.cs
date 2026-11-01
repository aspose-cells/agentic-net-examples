// Title: Convert an in‑memory HTML string to a PDF with embedded fonts using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML snippet from a string, loads it into a MemoryStream, creates an Aspose.Cells Workbook, and saves it as a PDF with fonts embedded. | Show how to configure PdfSaveOptions in Aspose.Cells to enable font embedding when converting HTML content to PDF. | Provide an example that converts HTML stored only in memory to a PDF file without writing the HTML to disk, using Aspose.Cells.
// Common Searches: asp.net convert html string to pdf with embedded fonts using aspose.cells | load html from memory stream into aspose.cells workbook c# | pdfsaveoptions embed fonts aspose.cells example | convert in‑memory html to pdf without saving html file c# | aspose.cells html to pdf conversion with font embedding settings
// Tags: Aspose.Cells HTML to PDF conversion | PdfSaveOptions embed fonts | MemoryStream HTML loading Aspose.Cells | C# generate PDF from HTML string | embedded fonts in PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads an HTML snippet into a UTF‑8 MemoryStream, creates an Aspose.Cells Workbook from the stream, and saves it as a PDF using PdfSaveOptions that embed the required fonts.
class Program
{
    static void Main()
    {
        try
        {
            // HTML content to be converted
            string html = "<html><body><h1>Sample Report</h1><p>This is a paragraph.</p></body></html>";

            // Convert the HTML string to a UTF‑8 byte array and wrap it in a MemoryStream
            using (MemoryStream htmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html)))
            {
                // Load the HTML stream into an Aspose.Cells Workbook
                Workbook workbook = new Workbook(htmlStream, new LoadOptions(LoadFormat.Html));

                // Configure PDF save options (default options are sufficient for this example)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF file
                workbook.Save("Result.pdf", pdfOptions);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
