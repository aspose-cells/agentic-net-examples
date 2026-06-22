using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    public class ExportPdfWithAttachments
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data to the worksheet
                worksheet.Cells["A1"].PutValue("PDF Export with Embedded Attachments and Optimization");

                // Create a temporary file that will be embedded as an OLE object
                string tempFilePath = Path.Combine(Path.GetTempPath(), "SampleAttachment.txt");
                File.WriteAllText(tempFilePath, "This is a sample attachment for PDF export.");

                // Ensure the temporary file exists before reading its bytes
                if (!File.Exists(tempFilePath))
                    throw new FileNotFoundException("Attachment file not found.", tempFilePath);

                // Add the OLE object (attachment) to the worksheet
                int oleIndex = worksheet.OleObjects.Add(5, 0, 200, 200, File.ReadAllBytes(tempFilePath));
                worksheet.OleObjects[oleIndex].DisplayAsIcon = true; // Show as icon

                // Configure PDF save options (only supported properties are used)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Embed OLE attachment in PDF
                };

                // Save the workbook as a PDF file with the specified options
                string outputPdf = "ExportedWithAttachments.pdf";
                workbook.Save(outputPdf, pdfOptions);

                // Clean up the temporary attachment file
                try
                {
                    if (File.Exists(tempFilePath))
                        File.Delete(tempFilePath);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Warning: Could not delete temporary file. {ex.Message}");
                }

                Console.WriteLine($"PDF saved to '{outputPdf}' with embedded attachment.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            ExportPdfWithAttachments.Run();
        }
    }
}