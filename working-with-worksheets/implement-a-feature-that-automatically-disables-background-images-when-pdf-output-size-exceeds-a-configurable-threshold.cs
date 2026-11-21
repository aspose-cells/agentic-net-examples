// Title: How to automatically disable worksheet background images when exporting to PDF if the file size exceeds a configurable limit using Aspose.Cells for .NET
// Common Searches: c# Aspose.Cells export to PDF and remove background images when file size is over 5 MB | conditionally clear worksheet background image before saving PDF with Aspose.Cells based on output size | measure PDF size in a MemoryStream and disable backgrounds in Aspose.Cells .NET | Aspose.Cells PDF generation size threshold background image removal example | auto disable Excel sheet background images for large PDF output using Aspose.Cells
// Tags: Aspose.Cells PDF size based background control | disable worksheet background image C# Aspose.Cells | evaluate PDF output size with MemoryStream Aspose.Cells | PdfSaveOptions re‑save after background removal | configurable size threshold PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, saves it to a PDF using PdfSaveOptions into a MemoryStream, checks the generated PDF size, and if it exceeds a 5 MB threshold it clears the BackgroundImage property of every worksheet and re‑saves the PDF before writing the final file to disk.
class PdfExportWithBackgroundControl
{
    // Configurable size threshold in bytes (e.g., 5 MB)
    const long SizeThresholdBytes = 5 * 1024 * 1024;

    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // First save to a memory stream to evaluate the output size
            using (MemoryStream tempStream = new MemoryStream())
            {
                workbook.Save(tempStream, pdfOptions);
                long outputSize = tempStream.Length;

                // If the generated PDF exceeds the threshold, remove background images and re‑save
                if (outputSize > SizeThresholdBytes)
                {
                    // Remove background images from all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        sheet.BackgroundImage = null;
                    }

                    // Reset the stream for the second save
                    tempStream.SetLength(0);
                    workbook.Save(tempStream, pdfOptions);
                }

                // Write the final PDF to disk
                File.WriteAllBytes(outputPath, tempStream.ToArray());
                Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
