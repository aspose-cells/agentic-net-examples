// Title: Auto‑remove worksheet background images when PDF exceeds size limit (Aspose.Cells C#)
// Description: Demonstrates a C# routine that saves a Workbook to PDF, checks the generated file size against a configurable byte threshold, and automatically clears all worksheet BackgroundImage properties before re‑saving if the size is too large. The approach keeps PDFs within attachment or bandwidth limits without manual intervention.
// Keywords: Aspose.Cells PDF size limit | C# remove worksheet background image | conditional PDF export Aspose.Cells | auto background removal PDF | Excel to PDF size check | Aspose.Cells memory stream | threshold based PDF generation
// Common Searches: Aspose.Cells disable background images for large PDF | C# export Excel to PDF only if under 5 MB | how to trim PDF size by removing worksheet backgrounds | auto remove sheet background when PDF too big Aspose | conditional PDF generation based on file size C#
// Developer Intent: Create a PDF from a workbook and automatically strip worksheet background images when the output exceeds a predefined size.
// Use Cases: Sending Excel‑derived reports via email where attachment size must stay below a limit. | Providing on‑demand PDFs in web apps that need to respect bandwidth caps. | Running batch conversions that adapt to storage quotas by removing heavy background graphics only when necessary.
// AI Prompts: Generate a C# method using Aspose.Cells that checks PDF size and removes all worksheet background images if it exceeds a configurable threshold. | Show how to read the size threshold from appsettings.json and apply it in the auto‑background‑removal workflow. | Explain logging of original and reduced PDF sizes and handling cases where no background images exist.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates a C# routine that saves a Workbook to PDF, checks the generated file size against a configurable byte threshold, and automatically clears all worksheet BackgroundImage properties before re‑saving if the size is too large. The approach keeps PDFs within attachment or bandwidth limits without manual intervention.
public class PdfExporter
{
    // Configurable size threshold in bytes (e.g., 5 MB)
    private const long SizeThresholdBytes = 5 * 1024 * 1024;

    /// <param name="workbook">The workbook to export.</param>
    /// <param name="outputPath">Full path of the resulting PDF file.</param>
    public static void SavePdfWithAutoBackgroundRemoval(Workbook workbook, string outputPath)
    {
        try
        {
            // Render PDF with default options into a memory stream
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            using (MemoryStream tempStream = new MemoryStream())
            {
                workbook.Save(tempStream, pdfOptions);

                // If size is within the acceptable range, write the stream to the final file
                if (tempStream.Length <= SizeThresholdBytes)
                {
                    File.WriteAllBytes(outputPath, tempStream.ToArray());
                    Console.WriteLine($"PDF saved successfully (size {tempStream.Length} bytes).");
                    return;
                }

                // Size exceeds threshold – remove background images from all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.BackgroundImage != null && sheet.BackgroundImage.Length > 0)
                    {
                        sheet.BackgroundImage = null;
                    }
                }

                // Regenerate PDF after background removal
                using (MemoryStream finalStream = new MemoryStream())
                {
                    workbook.Save(finalStream, pdfOptions);
                    File.WriteAllBytes(outputPath, finalStream.ToArray());
                    Console.WriteLine($"PDF size exceeded threshold. Background images removed. New size {finalStream.Length} bytes.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during PDF export: {ex.Message}");
        }
    }

    // Example usage
    public static void Main()
    {
        const string inputPath = "InputWithBackground.xlsx";
        const string outputPath = "Result.pdf";

        try
        {
            Workbook wb;

            if (File.Exists(inputPath))
            {
                // Load existing workbook
                wb = new Workbook(inputPath);
            }
            else
            {
                // Create a sample workbook if the input file is missing
                Console.WriteLine($"Input file '{inputPath}' not found. Creating a sample workbook.");
                wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Name = "SampleSheet";
                ws.Cells["A1"].PutValue("Sample Data");
            }

            SavePdfWithAutoBackgroundRemoval(wb, outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
