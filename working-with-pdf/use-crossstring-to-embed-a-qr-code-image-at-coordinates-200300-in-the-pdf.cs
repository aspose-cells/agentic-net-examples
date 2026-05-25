using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("QR Code embedded at (200,300)");

        // Load QR code image bytes (ensure "qr.png" exists in the executable directory)
        byte[] qrImageData = File.ReadAllBytes("qr.png");

        // Create an image watermark using the QR code bytes
        RenderingWatermark watermark = new RenderingWatermark(qrImageData)
        {
            // Position the watermark at the required coordinates (pixels)
            OffsetX = 200,          // X coordinate
            OffsetY = 300,          // Y coordinate
            // Align to the top‑left corner so offsets are absolute
            HAlignment = TextAlignmentType.Left,
            VAlignment = TextAlignmentType.Top,
            // Make the watermark fully opaque and keep original size
            Opacity = 1.0f,
            ScaleToPagePercent = 100,
            IsBackground = false   // Place in front of page content
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the QR code positioned at (200,300)
        workbook.Save("QrCodeEmbedded.pdf", pdfOptions);
    }
}