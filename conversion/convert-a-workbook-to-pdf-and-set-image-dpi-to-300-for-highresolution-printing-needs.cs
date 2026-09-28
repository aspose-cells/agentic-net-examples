// Title: Export an Excel workbook to PDF with 300 DPI images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures PdfSaveOptions to render images at 300 DPI (using reflection if the ImageDPI property is missing), and saves the workbook as a PDF with Aspose.Cells. | Show how to add file‑existence verification and comprehensive exception handling to an Aspose.Cells Excel‑to‑PDF conversion routine that requires high‑resolution images. | Demonstrate creating a print‑ready PDF from a workbook where charts and pictures are embedded at 300 DPI by leveraging PdfSaveOptions in Aspose.Cells.
// Common Searches: how to set image DPI to 300 when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells export workbook to high resolution PDF using PdfSaveOptions ImageDPI | C# reflection to set PdfSaveOptions.ImageDPI property only if it exists | check if input Excel file exists before saving as PDF with Aspose.Cells | exception handling pattern for Excel to PDF conversion using Aspose.Cells .NET
// Tags: PdfSaveOptions high DPI images | Aspose.Cells high‑resolution PDF export | C# reflection optional property | file existence validation before conversion | exception handling Aspose.Cells workflow

using Aspose.Cells;
using System;
using System.IO;

// The example checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates PdfSaveOptions, uses reflection to set the ImageDPI property to 300 DPI when the property is available, and saves the workbook as a high‑resolution PDF while handling any runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (high‑resolution images)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // The ImageDPI property may not be available in older versions of Aspose.Cells.
            // If it exists, set it to 300 DPI; otherwise, proceed without setting it.
            var imageDpiProperty = typeof(PdfSaveOptions).GetProperty("ImageDPI");
            if (imageDpiProperty != null && imageDpiProperty.CanWrite)
            {
                imageDpiProperty.SetValue(pdfOptions, 300);
            }

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
