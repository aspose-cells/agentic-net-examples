// Title: Convert an Excel workbook to a 600 DPI PDF with Aspose.Cells for .NET
// Description: C# example that loads an .xlsx file, sets the global rendering DPI to 600, configures PdfSaveOptions to resample images at 600 PPI with full quality, and saves a high‑resolution PDF ideal for detailed engineering drawings. Includes file‑existence verification and robust error handling.
// Keywords: Aspose.Cells | Excel to PDF | 600 DPI | high resolution PDF | PdfSaveOptions | SetImageResample | CellsHelper.DPI | engineering drawing export | C# conversion | PDF image quality
// Common Searches: Aspose.Cells export Excel to PDF 600 DPI | Set image DPI in PDF using Aspose.Cells .NET | High resolution PDF from Excel C# | PdfSaveOptions SetImageResample example | Increase PDF image quality Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook where all embedded images are rendered at 600 DPI for precise engineering drawings.
// Use Cases: Printing engineering schematics that require 600 DPI raster quality | Creating archival PDFs of technical spreadsheets | Batch processing design worksheets into high‑resolution PDFs | Delivering client‑ready reports with crisp graphics
// AI Prompts: Provide C# code that reads an .xlsx workbook and saves it as a PDF with 600 DPI image rendering using Aspose.Cells. | Explain how CellsHelper.DPI and PdfSaveOptions.SetImageResample affect PDF output quality in Aspose.Cells. | Show how to add file existence validation and exception handling to an Excel‑to‑PDF conversion script with high DPI settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // C# example that loads an .xlsx file, sets the global rendering DPI to 600, configures PdfSaveOptions to resample images at 600 PPI with full quality, and saves a high‑resolution PDF ideal for detailed engineering drawings. Includes file‑existence verification and robust error handling.
    public class ExportToPdfWithHighDpi
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
                // Set the global DPI to 600 for high‑resolution rendering
                CellsHelper.DPI = 600;

                const string inputPath = "EngineeringDrawing.xlsx";
                const string outputPath = "EngineeringDrawing_600dpi.pdf";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure images inside the PDF are rendered at 600 PPI and keep maximum quality
                pdfOptions.SetImageResample(600, 100);

                // Save the workbook as a PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
