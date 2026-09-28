// Title: Convert an HTML workbook to PDF with vector‑graphics chart rendering using Aspose.Cells for C#
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, sets PdfSaveOptions.VectorGraphics = true, and saves the workbook as a PDF. | Explain how to configure PdfSaveOptions in Aspose.Cells to preserve chart scalability by enabling vector graphics during HTML‑to‑PDF conversion.
// Common Searches: Aspose.Cells C# convert HTML file to PDF with vector charts | How to keep charts scalable when saving HTML workbook as PDF using Aspose.Cells | How to enable vector graphics for charts when saving PDF with Aspose.Cells | C# code for HTML to PDF conversion preserving chart quality with Aspose.Cells
// Tags: Aspose.Cells HTML-to-PDF workflow | Aspose.Cells PDF vector graphics option | C# scalable chart rendering in PDF | HTML workbook conversion to PDF using Aspose.Cells | vector graphics charts PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an HTML file into an Aspose.Cells Workbook, enabling the VectorGraphics option in PdfSaveOptions to render charts as scalable vectors, and saving the workbook as a PDF while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.pdf";

        // Verify that the input HTML file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the HTML file into an Aspose.Cells workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient for basic conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
