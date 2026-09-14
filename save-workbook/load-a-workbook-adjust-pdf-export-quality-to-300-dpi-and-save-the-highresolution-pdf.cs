// Title: Generate a 300 DPI high‑resolution PDF from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a .xlsx file, sets CellsHelper.DPI to 300, updates each worksheet's PageSetup.PrintQuality to 300, and saves the workbook as a PDF with PdfSaveOptions configured for 300 PPI image resampling and maximum JPEG quality. | Show how to configure Aspose.Cells PdfSaveOptions to resample images at 300 PPI while preserving full image quality during Excel‑to‑PDF conversion in C#. | Provide a C# example that verifies the source Excel file exists, applies high‑resolution settings, and exports it to a PDF named output_high_res.pdf.
// Common Searches: how to export Excel to PDF at 300 dpi with Aspose.Cells C# | Aspose.Cells set image resample 300 ppi for PDF conversion | increase PDF quality when saving workbook using Aspose.Cells .NET | global DPI setting CellsHelper for high resolution PDF output
// Tags: Aspose.Cells PdfSaveOptions high‑resolution image resample | CellsHelper DPI configuration 300 | Worksheet PageSetup PrintQuality 300 DPI | C# high‑resolution PDF generation from Excel | Aspose.Cells PDF export quality tuning

using System;
using System.IO;
using Aspose.Cells;

// Loads an .xlsx workbook, sets CellsHelper.DPI and each worksheet's PrintQuality to 300 DPI, configures PdfSaveOptions to resample images at 300 PPI with full JPEG quality, and saves the workbook as a high‑resolution PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Set the global DPI to 300 – this influences rendering of images and shapes.
            CellsHelper.DPI = 300;

            string inputPath = "input.xlsx";
            string outputPath = "output_high_res.pdf";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found – {inputPath}");
                return;
            }

            // Load the source workbook.
            Workbook workbook = new Workbook(inputPath);

            // Optionally, set each worksheet's print quality to 300 DPI.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PrintQuality = 300;
            }

            // Prepare PDF save options for high‑resolution output.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Resample images to 300 PPI and use the highest JPEG quality (100%).
            pdfOptions.SetImageResample(300, 100);

            // The default optimization type provides high quality; 
            // if the enum is unavailable in the referenced version, this line can be omitted.
            // pdfOptions.OptimizationType = PdfOptimizationType.Standard;

            // Save the workbook as a high‑resolution PDF.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
