using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfCompliance enum

// Author: Aspose.Cells .NET example – verify PDF save options
class PdfOptionsVerification
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample workbook with minimal content
        // ------------------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Verification Test");

        // ------------------------------------------------------------
        // 2. Configure PDF save options
        // ------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set PDF/A-1b compliance
            Compliance = PdfCompliance.PdfA1b,

            // Explicitly set the creation time (will be written into PDF metadata)
            CreatedTime = new DateTime(2023, 1, 1, 12, 0, 0),

            // Set a custom producer string
            Producer = "PdfOptionsVerificationDemo",

            // Force each worksheet onto a single page
            OnePagePerSheet = true
        };

        // ------------------------------------------------------------
        // 3. Save the workbook as PDF using the configured options
        // ------------------------------------------------------------
        const string pdfPath = "VerificationOutput.pdf";
        workbook.Save(pdfPath, pdfOptions);

        // ------------------------------------------------------------
        // 4. Basic file‑system verification (exists, non‑empty)
        // ------------------------------------------------------------
        FileInfo fileInfo = new FileInfo(pdfPath);
        if (!fileInfo.Exists)
        {
            Console.WriteLine("Error: PDF file was not created.");
            return;
        }

        if (fileInfo.Length == 0)
        {
            Console.WriteLine("Error: PDF file is empty.");
            return;
        }

        Console.WriteLine($"PDF created successfully. Size: {fileInfo.Length} bytes.");

        // ------------------------------------------------------------
        // 5. Verify PDF metadata (producer, creation time, compliance)
        // ------------------------------------------------------------
        // NOTE: Aspose.Pdf is typically used to read PDF metadata.
        // The following code assumes Aspose.Pdf is referenced.
        // If Aspose.Pdf is not available, replace with appropriate library
        // or inspect the file manually.

        try
        {
            // Placeholder for actual PDF metadata extraction
            // using Aspose.Pdf;
            // PdfDocument pdfDoc = new PdfDocument(pdfPath);
            // string producer = pdfDoc.Info.Producer;
            // DateTime? created = pdfDoc.Info.CreationDate;
            // Console.WriteLine($"Producer: {producer}");
            // Console.WriteLine($"Creation Time (PDF metadata): {created}");

            // Since Aspose.Pdf API is not documented here, we provide a stub:
            Console.WriteLine("Metadata verification requires Aspose.Pdf or another PDF library.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Metadata verification failed: {ex.Message}");
        }

        // ------------------------------------------------------------
        // 6. Additional runtime checks (optional)
        // ------------------------------------------------------------
        // Example: ensure the file's last write time matches the set CreatedTime
        // (File system timestamps may differ from PDF internal timestamps.)
        DateTime fileWriteTime = fileInfo.LastWriteTime;
        Console.WriteLine($"File system last write time: {fileWriteTime}");

        // End of verification
    }
}