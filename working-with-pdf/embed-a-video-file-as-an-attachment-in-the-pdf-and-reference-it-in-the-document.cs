// Title: Embed MP4 Video as a PDF Attachment with Aspose.Cells (C#)
// Description: Demonstrates how to read an MP4 file, add an OLE object placeholder in a workbook, embed the video bytes, enable attachment embedding via PdfSaveOptions, and export the workbook to a PDF that contains the video as an attached file.
// Keywords: Aspose.Cells embed video PDF | C# OLE object video attachment | PdfSaveOptions EmbedAttachments | MP4 attachment Aspose.Cells | export Excel to PDF with video
// Common Searches: Aspose.Cells add video attachment to PDF | C# embed MP4 in PDF using Aspose.Cells | How to use PdfSaveOptions to embed files | Create PDF with embedded video from Excel | OLE object video export Aspose.Cells .NET
// Developer Intent: Add a video file as an embedded attachment in a PDF generated from an Aspose.Cells workbook.
// Use Cases: Product catalog PDFs where each item links to a demo video. | Training manuals that include tutorial videos accessible from the worksheet. | Sales presentations with promotional videos attached and shown as icons.
// AI Prompts: Provide C# code to embed an MP4 as an OLE object in an Aspose.Cells workbook and save it as a PDF with the video attached. | Explain how to configure PdfSaveOptions to embed attachments when exporting a workbook to PDF with Aspose.Cells. | Show an example of adding multiple video attachments to a worksheet and exporting them into a single PDF file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsVideoAttachmentDemo
{
    // Demonstrates how to read an MP4 file, add an OLE object placeholder in a workbook, embed the video bytes, enable attachment embedding via PdfSaveOptions, and export the workbook to a PDF that contains the video as an attached file.
    class Program
    {
        static void Main()
        {
            // Path to the video file to embed
            string videoPath = "sample.mp4";

            // Verify video file exists
            if (!File.Exists(videoPath))
            {
                Console.WriteLine($"Video file not found: {videoPath}");
                return;
            }

            // Read video bytes
            byte[] videoBytes = File.ReadAllBytes(videoPath);

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a label in the worksheet
            sheet.Cells["A1"].PutValue("Embedded Video Attachment");

            // Add an OLE object placeholder (empty image data)
            int oleIndex = sheet.OleObjects.Add(5, 1, 200, 200, new byte[0]);

            // Get the OLE object reference
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Embed the video file into the OLE object
            // Parameters: linkToFile = false (embed), objectData = videoBytes,
            // sourceFileName = "sample.mp4", displayAsIcon = true, label = "Video",
            // updateIcon = false (keep default icon)
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: videoBytes,
                sourceFileName: "sample.mp4",
                displayAsIcon: true,
                label: "Video",
                updateIcon: false);

            // Configure PDF save options to embed attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as PDF with the embedded video attachment
            string outputPdf = "WorkbookWithVideo.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved successfully: {outputPdf}");
        }
    }
}
