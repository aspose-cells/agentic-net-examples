// Title: Batch add a consistent watermark while converting Excel to PDF with Aspose.Cells for .NET (C#)
// Description: Creates input/output folders, defines a reusable RenderingWatermark (centered, rotated, semi‑transparent), loops through all *.xlsx files, applies the watermark via PdfSaveOptions, and saves each workbook as a watermarked PDF.
// Keywords: Aspose.Cells watermark PDF C# | batch Excel to PDF conversion | RenderingWatermark example | reusable watermark Aspose.Cells | automate PDF watermarking
// Common Searches: apply same watermark to multiple Excel files Aspose.Cells | C# batch convert .xlsx to PDF with diagonal watermark | loop RenderingWatermark with PdfSaveOptions | reuse watermark instance for many workbooks | watermark all PDFs generated from Excel folder
// Developer Intent: Generate watermarked PDFs from every Excel workbook in a directory using a single shared watermark definition.
// Use Cases: Produce confidential PDFs for a batch of financial reports. | Automate watermarking of contract spreadsheets before client delivery. | Create a searchable, watermarked archive of daily sales Excel files.
// AI Prompts: Give C# code that adds a semi‑transparent diagonal watermark to every worksheet in a workbook with Aspose.Cells. | Show how to extend the batch processor to handle .xls files and set a custom watermark text per file. | Explain how to vary watermark opacity and rotation based on workbook metadata at runtime.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates input/output folders, defines a reusable RenderingWatermark (centered, rotated, semi‑transparent), loops through all *.xlsx files, applies the watermark via PdfSaveOptions, and saves each workbook as a watermarked PDF.
class BatchWatermarkProcessor
{
    static void Main()
    {
        // Directory containing source Excel files
        string inputDir = @"C:\InputExcel";

        // Directory where watermarked PDFs will be saved
        string outputDir = @"C:\OutputPdf";
        Directory.CreateDirectory(outputDir);

        // Define a consistent watermark font
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Italic = true,
            Color = Color.Red
        };

        // Create a single watermark instance to reuse for all files
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,
            Opacity = 0.3f,
            ScaleToPagePercent = 75,
            IsBackground = true
        };

        // Process each .xlsx file in the input directory
        foreach (string filePath in Directory.GetFiles(inputDir, "*.xlsx"))
        {
            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Determine output PDF path
            string outputFile = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

            // Save the workbook as PDF with the watermark applied
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"Watermarked PDF saved: {outputFile}");
        }
    }
}
