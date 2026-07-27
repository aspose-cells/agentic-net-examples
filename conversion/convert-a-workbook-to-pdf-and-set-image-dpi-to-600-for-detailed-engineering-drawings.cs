// Title: Export Excel Workbook to PDF with 600 DPI Images Using Aspose.Cells for .NET
// Description: Loads an Excel file, sets the global DPI to 600, configures PdfSaveOptions to resample images at 600 PPI with maximum JPEG quality, and saves the workbook as a high‑resolution PDF suitable for detailed engineering drawings.
// Keywords: Aspose.Cells | PDF conversion .NET | 600 DPI | high resolution PDF | PdfSaveOptions SetImageResample | engineering drawing export | C# Excel to PDF | image quality control | raster graphics DPI | Aspose.Cells DPI setting
// Common Searches: Aspose.Cells export Excel to PDF 600 DPI | set image resolution when saving PDF with Aspose.Cells | high‑resolution PDF from workbook .NET | PdfSaveOptions SetImageResample example | increase PDF image quality Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook where every embedded image is rendered at 600 DPI to preserve the fidelity of engineering drawings.
// Use Cases: Generate printable engineering schematics from Excel templates with raster graphics at 600 PPI. | Batch‑process design spreadsheets into archival‑grade PDFs that meet ISO drawing standards. | Integrate high‑resolution PDF export into CI/CD pipelines for automated documentation generation.
// AI Prompts: How can I make the DPI value configurable at runtime for the PDF export? | Show an example that keeps vector shapes sharp while setting image DPI to 600. | What strategies help reduce memory usage when converting large workbooks to high‑resolution PDFs with Aspose.Cells?

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

namespace EngineeringDrawingExport
{
    // Loads an Excel file, sets the global DPI to 600, configures PdfSaveOptions to resample images at 600 PPI with maximum JPEG quality, and saves the workbook as a high‑resolution PDF suitable for detailed engineering drawings.
    public class WorkbookToPdf
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Set the machine DPI to 600. This influences image rendering quality.
                CellsHelper.DPI = 600;

                // Path to the source workbook
                string sourcePath = "EngineeringDrawing.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Error: Source file not found at path '{sourcePath}'.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(sourcePath);

                // Configure PDF save options to ensure images are kept at 600 PPI
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                // Desired PPI = 600, JPEG quality = 100 (maximum)
                pdfOptions.SetImageResample(600, 100);

                // Path for the output PDF
                string outputPath = "EngineeringDrawing.pdf";

                // Save the workbook as a PDF with the specified DPI settings
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully converted to PDF with 600 DPI at: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
