// Title: Embedding a video file as an attachment in a PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add MP4 video as PDF attachment | how to create a hyperlink to a video attachment in a PDF using Aspose.Cells | Aspose.Cells PDFSaveOptions embed media file limitation | reference an attached video file from a cell hyperlink in a generated PDF | why Aspose.Cells does not embed video files when saving to PDF
// Tags: Aspose.Cells PDF attachment video | C# PDF attachment URI scheme Aspose.Cells | hyperlink to embedded media in PDF workbook | Aspose.Cells PDFSaveOptions attachment support | embedding media files in PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, writes link text in cell A1, adds a hyperlink using the `attachment://sample_video.mp4` URI scheme, and saves the workbook as a PDF with `PdfSaveOptions`. Aspose.Cells for .NET does not currently embed the video file as an attachment in the generated PDF, so the hyperlink references a non‑existent attachment.
class EmbedVideoInPdf
{
    static void Main()
    {
        try
        {
            // Paths – adjust as needed
            string videoPath = @"C:\Path\To\Your\sample_video.mp4";
            string outputPath = @"C:\Path\To\Output\WorkbookWithVideo.pdf";

            // Verify the video file exists
            if (!File.Exists(videoPath))
            {
                Console.WriteLine($"Video file not found: {videoPath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Write link text
            sheet.Cells["A1"].PutValue("Click here to open the video attachment");

            // Add a hyperlink that points to the attachment using the PDF attachment URI scheme
            sheet.Hyperlinks.Add(0, 0, 0, 0, "attachment://sample_video.mp4");

            // Prepare PDF save options (no attachment support in Aspose.Cells PDF conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
