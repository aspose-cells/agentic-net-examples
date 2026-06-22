using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAttachmentDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("PDF with Text File Attachment");

                // Prepare a temporary text file to be attached
                string tempTxtPath = Path.Combine(Path.GetTempPath(), "Attachment.txt");
                File.WriteAllText(tempTxtPath, "This is the content of the attached text file.");

                // Ensure the temporary file exists before reading it
                if (!File.Exists(tempTxtPath))
                    throw new FileNotFoundException("Temporary attachment file not found.", tempTxtPath);

                // Add the text file as an OLE object (attachment) to the worksheet
                // Row, Column, Width, Height are arbitrary values for placement
                int oleIndex = sheet.OleObjects.Add(5, 0, 200, 200, File.ReadAllBytes(tempTxtPath));

                // Set the OLE object to display as an icon
                sheet.OleObjects[oleIndex].DisplayAsIcon = true;

                // (Optional) Set file format type – using Unknown if specific type is unavailable
                sheet.OleObjects[oleIndex].FileFormatType = FileFormatType.Unknown;

                // Configure PDF save options to embed OLE attachments
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Enables embedding of the OLE object into the PDF
                };

                // Save the workbook as PDF with the attachment embedded
                string outputPdf = "WorkbookWithAttachment.pdf";
                workbook.Save(outputPdf, pdfOptions);

                // Clean up the temporary text file
                if (File.Exists(tempTxtPath))
                {
                    File.Delete(tempTxtPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}