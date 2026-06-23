using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EmbedTextAsPdfAttachment
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Wrap the whole operation in a try‑catch to handle unexpected errors.
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("PDF with Embedded Text Attachment Example");

                // Create a temporary text file to embed.
                string txtFilePath = "sample.txt";
                File.WriteAllText(txtFilePath, "This is the content of the embedded text document.");

                // Ensure the file exists before reading its bytes.
                if (!File.Exists(txtFilePath))
                    throw new FileNotFoundException("Temporary text file was not created.", txtFilePath);

                // Add the text file as an OLE object (attachment) to the worksheet.
                // Parameters: upper‑left row, upper‑left column, height, width, file bytes.
                int oleIndex = worksheet.OleObjects.Add(10, 10, 200, 200, File.ReadAllBytes(txtFilePath));

                // Show the attachment as an icon.
                worksheet.OleObjects[oleIndex].DisplayAsIcon = true;

                // Configure PDF save options to embed OLE attachments.
                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true
                };

                // Save the workbook as a PDF file with the embedded text attachment.
                string pdfPath = "PdfWithEmbeddedTextAttachment.pdf";
                workbook.Save(pdfPath, pdfSaveOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary text file if it exists.
                string txtFilePath = "sample.txt";
                if (File.Exists(txtFilePath))
                {
                    try
                    {
                        File.Delete(txtFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                    }
                }
            }
        }
    }
}