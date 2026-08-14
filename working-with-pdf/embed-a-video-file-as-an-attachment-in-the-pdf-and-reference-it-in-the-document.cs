// Title: Embed a Video File as an OLE Attachment in a PDF with Aspose.Cells (C#)
// Description: This example creates a workbook, adds a text label, reads an MP4 file into a byte array, inserts an OLE placeholder, embeds the video bytes as an OLE object with an icon, enables the EmbedAttachments flag in PdfSaveOptions, and saves the workbook as a PDF that contains the video as an attachment.
// Keywords: Aspose.Cells | C# | embed video PDF | OLE object | PdfSaveOptions | EmbedAttachments | MP4 attachment | PDF generation | Aspose.Cells tutorial
// Common Searches: Aspose.Cells embed video in PDF C# | How to add MP4 as attachment using Aspose.Cells | SetEmbeddedObject video OLE Aspose.Cells | Enable EmbedAttachments in PdfSaveOptions | C# code to attach video to PDF with Aspose
// Developer Intent: The developer wants to attach a video file to a PDF generated from an Aspose.Cells workbook and make the attachment accessible from the document.
// Use Cases: Product catalogs where each item includes a demonstration video embedded in the PDF. | Training manuals that provide video tutorials linked to specific worksheet sections. | Compliance reports that need to bundle video evidence as PDF attachments.
// AI Prompts: Generate C# code that embeds multiple MP4 files as OLE attachments in one PDF using Aspose.Cells, with robust error handling for missing files. | Explain how to extract and play a video attachment from a PDF created with Aspose.Cells' EmbedAttachments option. | Show how to customize the icon and label of an embedded video OLE object before saving the workbook to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsVideoAttachmentDemo
{
    // This example creates a workbook, adds a text label, reads an MP4 file into a byte array, inserts an OLE placeholder, embeds the video bytes as an OLE object with an icon, enables the EmbedAttachments flag in PdfSaveOptions, and saves the workbook as a PDF that contains the video as an attachment.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a descriptive text in the worksheet
            worksheet.Cells["A1"].PutValue("Embedded Video Attachment Demo");

            // Path to the video file to embed (ensure the file exists)
            string videoPath = "sample_video.mp4";

            if (!File.Exists(videoPath))
            {
                Console.WriteLine($"Video file not found: {videoPath}");
                return;
            }

            // Read video file bytes
            byte[] videoBytes = File.ReadAllBytes(videoPath);

            // Add an OLE object placeholder (empty image data) to the worksheet
            // The placeholder size is 200x200 pixels at row 5, column 2
            int oleIndex = worksheet.OleObjects.Add(5, 2, 200, 200, new byte[0]);

            // Get the added OleObject
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the video file into the OLE object
            // linkToFile = false (embed the data), displayAsIcon = true, label = "Video"
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: videoBytes,
                sourceFileName: Path.GetFileName(videoPath),
                displayAsIcon: true,
                label: "Sample Video");

            // Configure PDF save options to embed attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true // Enable embedding of OLE attachments in the PDF
            };

            // Save the workbook as PDF; the video will be embedded as an attachment
            string outputPdf = "Workbook_With_Video.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved with embedded video: {outputPdf}");
        }
    }
}
