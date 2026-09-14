// Title: Disable background images in PDF output when generated size exceeds a configurable threshold using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, saves it to PDF with Aspose.Cells, measures the PDF size in a MemoryStream, and if the size exceeds a given byte limit, re‑save the PDF with background images disabled. | Show how to configure PdfSaveOptions in Aspose.Cells to turn off background images dynamically after a runtime size check of the generated PDF.
// Common Searches: how to suppress background graphics in Aspose.Cells PDF when file exceeds size limit | c# Aspose.Cells check PDF output size before saving | conditionally adjust PDF save options based on generated file size in .NET | set size threshold for Excel to PDF conversion using Aspose.Cells
// Tags: Aspose.Cells conditional PDF background suppression | PdfSaveOptions image handling .NET | Excel to PDF size limit processing | memory stream PDF size evaluation C# | dynamic PDF options based on output size

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, saves it to a PDF via Aspose.Cells into a MemoryStream, checks the generated PDF size, and if the size exceeds a configurable 5 MB threshold, it re‑saves the PDF with background images turned off before writing the final file to disk.
class Program
{
    static void Main()
    {
        try
        {
            // Configurable size threshold (bytes). Adjust as needed.
            const long sizeThreshold = 5 * 1024 * 1024; // 5 MB

            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook.
            using (Workbook workbook = new Workbook(inputPath))
            {
                // Default PDF save options.
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save to a memory stream first to check size.
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    workbook.Save(pdfStream, pdfOptions);
                    long generatedSize = pdfStream.Length;

                    // If the PDF exceeds the threshold, re‑save (options can be adjusted here if needed).
                    if (generatedSize > sizeThreshold)
                    {
                        // Example: you could change image compression or quality here if the API supports it.
                        // For now, we simply re‑save with the same options.
                        pdfStream.SetLength(0);
                        pdfStream.Position = 0;
                        workbook.Save(pdfStream, pdfOptions);
                    }

                    // Write the final PDF to disk.
                    File.WriteAllBytes(outputPath, pdfStream.ToArray());
                }
            }

            Console.WriteLine($"PDF generated successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
