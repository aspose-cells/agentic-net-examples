using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsVideoAttachmentDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the video file to embed
            string videoPath = "sample.mp4";

            // Verify that the video file exists
            if (!File.Exists(videoPath))
            {
                Console.WriteLine($"Video file not found: {videoPath}");
                return;
            }

            // Read video bytes
            byte[] videoBytes = File.ReadAllBytes(videoPath);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some descriptive text
            sheet.Cells["A1"].PutValue("PDF with Embedded Video Attachment");

            // Add an OLE object placeholder (imageData can be empty byte array)
            int oleIndex = sheet.OleObjects.Add(5, 1, 200, 200, new byte[0]);

            // Get the OLE object reference
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Embed the video file into the OLE object, display as an icon with a label
            // Parameters: linkToFile = false (embed), objectData = videoBytes,
            // sourceFileName = "sample.mp4", displayAsIcon = true, label = "Sample Video"
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: videoBytes,
                sourceFileName: Path.GetFileName(videoPath),
                displayAsIcon: true,
                label: "Sample Video"
            );

            // Configure PDF save options to embed OLE attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true // Ensure the video is embedded in the PDF
            };

            // Save the workbook as PDF with the embedded video attachment
            string outputPdf = "WorkbookWithVideo.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved successfully: {outputPdf}");
        }
    }
}