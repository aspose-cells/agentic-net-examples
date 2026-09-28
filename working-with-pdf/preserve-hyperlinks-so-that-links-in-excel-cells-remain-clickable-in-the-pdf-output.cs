// Title: Create a PDF from an Excel workbook that retains clickable hyperlinks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as a PDF while keeping all cell hyperlinks functional. | Show how to configure PdfSaveOptions in Aspose.Cells to ensure hyperlink preservation during Excel‑to‑PDF conversion. | Explain steps to verify that hyperlinks remain active in the generated PDF after using Aspose.Cells.
// Common Searches: how to keep Excel cell links active when converting to PDF with Aspose.Cells C# | Aspose.Cells PDF conversion preserving hyperlinks example | C# save workbook as PDF with clickable URLs using PdfSaveOptions | retain hyperlink functionality in PDF generated from .xlsx using Aspose | Aspose.Cells PDFSaveOptions default hyperlink behavior
// Tags: Aspose.Cells PDF hyperlink preservation | C# Excel to PDF active links | PdfSaveOptions keep hyperlinks | convert .xlsx to PDF clickable URLs | preserve cell hyperlinks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, creates a PdfSaveOptions object (hyperlinks are preserved by default), and saves the file as a PDF, ensuring that any hyperlinks embedded in the original cells stay clickable in the resulting document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing Excel file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (hyperlinks are preserved by default)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
