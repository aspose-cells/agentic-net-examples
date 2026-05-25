using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsAttachmentVerification
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare a temporary directory and attachment file
                string tempDir = Path.Combine(Path.GetTempPath(), "AsposeAttachmentDemo");
                Directory.CreateDirectory(tempDir);

                string attachmentPath = Path.Combine(tempDir, "sample.txt");
                File.WriteAllText(attachmentPath, "This is a sample attachment file.");

                // Verify the attachment file exists before using it
                if (!File.Exists(attachmentPath))
                    throw new FileNotFoundException("Attachment file not found.", attachmentPath);

                // Create a new workbook and add some content
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("PDF with Embedded Attachment");

                // Add the attachment as an OLE object (embed = true)
                byte[] attachmentBytes = File.ReadAllBytes(attachmentPath);
                int oleIndex = sheet.OleObjects.Add(5, 0, 200, 200, attachmentBytes);
                OleObject ole = sheet.OleObjects[oleIndex];
                // Use Unknown format for generic text file (Txt is not a valid enum value)
                ole.FileFormatType = FileFormatType.Unknown;
                ole.DisplayAsIcon = true;                // Show as an icon in the worksheet
                ole.Label = "Sample Text File";

                // Configure PDF save options to embed attachments
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Enable embedding of OLE attachments into the PDF
                };

                // Save the workbook as PDF
                string pdfPath = Path.Combine(tempDir, "WorkbookWithAttachment.pdf");
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF saved to: {pdfPath}");

                // NOTE: Verification of embedded attachments in the PDF requires Aspose.Pdf.
                // If Aspose.Pdf is available, you can load the PDF and inspect its EmbeddedFiles collection.
                // The verification code has been omitted to keep this project free of Aspose.Pdf dependencies.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}