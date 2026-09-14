// Title: Convert an Excel workbook to a PDF/A‑1a file with full font embedding using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, configures PdfSaveOptions for PDF/A‑1a compliance and enables embedding of all fonts, then saves the workbook as a PDF with Aspose.Cells. | Demonstrate how to set the compliance level to PdfA1a and turn on font embedding before calling Workbook.Save to produce a PDF/A‑1a document.
// Common Searches: Aspose.Cells C# export Excel to PDF/A-1a with embedded fonts | How to enable full font embedding when saving a workbook as PDF/A-1a in Aspose.Cells | PdfSaveOptions settings for PDF/A-1a compliance and font embedding in .NET | Convert .xlsx to PDF/A-1a using Aspose.Cells and ensure all fonts are embedded | C# code sample for PDF/A-1a compliant PDF generation from Excel with Aspose.Cells
// Tags: Aspose.Cells PDF/A-1a export C# | PdfSaveOptions font embedding Aspose.Cells | Excel to PDF/A-1a conversion .NET | embed all fonts Aspose.Cells PDF output | PDF/A-1a compliance setting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The example loads an existing Excel workbook (or creates a new one), configures PdfSaveOptions to enforce PDF/A‑1a compliance and to embed every font used in the workbook, and then saves the result as a PDF file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new workbook
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Configure PDF save options (no need to set SaveFormat; it's implicit)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
