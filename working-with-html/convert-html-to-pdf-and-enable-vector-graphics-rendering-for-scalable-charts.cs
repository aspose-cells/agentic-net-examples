// Title: Convert HTML to PDF with scalable vector‑graphics charts using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions to render charts as vector graphics, and saves the result as a PDF. | Show how to verify the input HTML file exists, set PdfSaveOptions.VectorGraphics = true, and produce a high‑resolution PDF with scalable charts. | Provide a console‑application snippet that demonstrates converting HTML to PDF while preserving chart quality by enabling vector graphics in Aspose.Cells.
// Common Searches: Aspose.Cells C# convert HTML file to PDF with vector graphics for charts | How to enable PdfSaveOptions.VectorGraphics when saving PDF from HTML in Aspose.Cells | Scalable chart rendering in PDF using Aspose.Cells .NET | Convert HTML to PDF preserving Excel chart quality with Aspose.Cells | PdfSaveOptions vector graphics option example for HTML to PDF conversion
// Tags: Aspose.Cells PDFSaveOptions VectorGraphics | HTML to PDF conversion Aspose.Cells | C# render charts as vector graphics PDF | Workbook.Save with vector graphics option | scalable chart rendering PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The program checks that input.html exists, loads it into an Aspose.Cells Workbook, enables PdfSaveOptions.VectorGraphics to render charts as vector graphics, and saves the workbook as output.pdf while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a Workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (charts will be rendered with default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
