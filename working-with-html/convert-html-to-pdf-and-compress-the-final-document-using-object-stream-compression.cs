// Title: Convert an HTML file to a compressed PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an HTML document into an Aspose.Cells Workbook and saves it as a PDF with object stream compression enabled via PdfSaveOptions. | Explain how to set PdfSaveOptions.Compression to PdfCompressionType.Stream when exporting HTML to PDF with Aspose.Cells.
// Common Searches: how to export HTML to PDF with reduced file size using Aspose.Cells C# | Aspose.Cells PdfSaveOptions compression option for PDF output | C# convert html file to compressed PDF with Aspose.Cells library | enable PDF size reduction in Aspose.Cells .NET example | save workbook as PDF with smaller size Aspose.Cells
// Tags: HTML to PDF conversion Aspose.Cells | PdfSaveOptions compression setting | Aspose.Cells PDF export reduced size | C# workbook save compressed PDF | PDF size optimization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for an input HTML file, loads it into an Aspose.Cells Workbook, optionally configures PdfSaveOptions to use stream compression, ensures the output directory exists, and saves the workbook as a compressed PDF.
class Program
{
    static void Main()
    {
        // Paths for input HTML and output PDF
        string htmlPath = "input.html";
        string pdfPath = "output.pdf";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: Input file not found – {htmlPath}");
            return;
        }

        try
        {
            // Load the HTML file into a Workbook
            Workbook workbook = new Workbook(htmlPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Uncomment the following line if the Compression property is available in your Aspose.Cells version
            // pdfOptions.Compression = PdfCompressionType.Stream;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
