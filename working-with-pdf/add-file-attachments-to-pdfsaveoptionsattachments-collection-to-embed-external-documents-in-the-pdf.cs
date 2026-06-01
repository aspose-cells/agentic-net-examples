using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

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
                Worksheet worksheet = workbook.Worksheets[0];

                // Put a title in the worksheet
                worksheet.Cells["A1"].PutValue("PDF with Embedded Attachments Example");

                // Create a temporary file that will be attached to the PDF
                string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleAttachment.txt");
                File.WriteAllText(tempFilePath, "This is the content of the attached file.");

                // Ensure the temporary file exists before reading its bytes
                if (!File.Exists(tempFilePath))
                    throw new FileNotFoundException("Attachment file not found.", tempFilePath);

                // Add the file as an OLE object (attachment) to the worksheet
                // Parameters: upper left row, upper left column, width, height, file bytes
                int oleIndex = worksheet.OleObjects.Add(10, 10, 200, 200, File.ReadAllBytes(tempFilePath));

                // Display the attachment as an icon
                worksheet.OleObjects[oleIndex].DisplayAsIcon = true;

                // Create PDF save options and enable embedding of attachments
                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true
                };

                // Save the workbook as PDF; the OLE object will be embedded as an attachment
                string outputPdf = Path.Combine(Path.GetTempPath(), "PdfWithEmbeddedAttachments.pdf");
                workbook.Save(outputPdf, pdfSaveOptions);

                Console.WriteLine($"PDF saved with embedded attachment: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary attachment file if it exists
                string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleAttachment.txt");
                if (File.Exists(tempFilePath))
                {
                    try
                    {
                        File.Delete(tempFilePath);
                    }
                    catch
                    {
                        // Suppress any exception during cleanup
                    }
                }
            }
        }
    }
}