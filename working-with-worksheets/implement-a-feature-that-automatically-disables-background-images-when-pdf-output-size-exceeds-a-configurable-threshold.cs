// Title: Aspose.Cells for .NET – Auto‑disable worksheet background images when PDF exceeds size limit
// Description: Demonstrates how to export a workbook to PDF, measure the generated file size, and automatically remove all worksheet background images if the size surpasses a configurable threshold (e.g., 5 MB). The example uses PdfSaveOptions with MinimumSize optimization, a temporary MemoryStream for size checking, and re‑saves the PDF after background removal.
// Keywords: Aspose.Cells PDF size limit | disable worksheet background image | conditional PDF export | C# Aspose.Cells PDF optimization | check PDF file size before save | remove background images Aspose.Cells | auto background removal PDF | minimum size PDF Aspose
// Common Searches: Aspose.Cells remove background image if PDF too large | C# export workbook to PDF with size threshold | how to limit PDF file size in Aspose.Cells | conditional background image removal during PDF export | auto disable worksheet background for large PDFs
// Developer Intent: Automatically drop worksheet background images when the exported PDF exceeds a predefined size limit.
// Use Cases: Generate compact PDF reports for email attachments by stripping backgrounds only when necessary. | Enforce file‑size policies in automated batch conversions of Excel workbooks to PDF. | Provide a fallback PDF version without backgrounds for low‑bandwidth environments.
// AI Prompts: Create a reusable C# method that takes a Workbook and a size limit, saves the PDF with backgrounds if under the limit, otherwise clears all BackgroundImage properties and re‑saves. | Show how to externalize the size threshold to appsettings.json and integrate the conditional background removal into an existing Aspose.Cells PDF export pipeline. | Write a GitHub‑style README snippet explaining the memory‑stream size check and background‑image removal logic for Aspose.Cells PDF generation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to export a workbook to PDF, measure the generated file size, and automatically remove all worksheet background images if the size surpasses a configurable threshold (e.g., 5 MB). The example uses PdfSaveOptions with MinimumSize optimization, a temporary MemoryStream for size checking, and re‑saves the PDF after background removal.
class PdfExportWithBackgroundControl
{
    static void Main()
    {
        try
        {
            // Configurable size threshold (e.g., 5 MB)
            const long sizeThresholdBytes = 5 * 1024 * 1024;

            // ---------- Create or load workbook ----------
            Workbook workbook = new Workbook(); // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF export.");

            // Add a background image for demonstration (optional)
            // Note: System.Drawing may not be available on all platforms, so this step is skipped.
            // If needed, ensure System.Drawing.Common is referenced and uncomment the code below.
            /*
            string backgroundImagePath = "background.jpg";
            if (File.Exists(backgroundImagePath))
            {
                sheet.BackgroundImage = System.Drawing.Image.FromFile(backgroundImagePath);
            }
            */

            // ---------- Prepare PDF save options ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use minimum size optimization as a first step
                OptimizationType = PdfOptimizationType.MinimumSize
            };

            // ---------- First save to memory to check size ----------
            using (MemoryStream tempStream = new MemoryStream())
            {
                workbook.Save(tempStream, pdfOptions);
                long generatedSize = tempStream.Length;

                // Determine output file path
                string outputPath = "output.pdf";

                // If size exceeds the threshold, remove background images and re‑save
                if (generatedSize > sizeThresholdBytes)
                {
                    // Disable background images on all worksheets
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        ws.BackgroundImage = null;
                    }

                    // Re‑save the workbook after background removal
                    using (FileStream finalFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Save(finalFile, pdfOptions);
                    }

                    Console.WriteLine($"PDF size ({generatedSize} bytes) exceeded threshold. Background images removed and PDF saved.");
                }
                else
                {
                    // Size is acceptable; save the original PDF
                    using (FileStream finalFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Save(finalFile, pdfOptions);
                    }

                    Console.WriteLine($"PDF size ({generatedSize} bytes) within threshold. PDF saved with background images.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
