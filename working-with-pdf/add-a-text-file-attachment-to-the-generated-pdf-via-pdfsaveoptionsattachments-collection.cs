// Title: C# – Add a Text File Attachment to a PDF Generated with Aspose.Cells (PdfSaveOptions.Attachments)
// Description: Demonstrates how to create an Aspose.Cells workbook, generate a temporary text file, embed it as an OLE object, enable PdfSaveOptions.EmbedAttachments, and save the workbook as a PDF that contains the text file attachment. The example also shows how to clean up the temporary file after saving.
// Keywords: Aspose.Cells PDF attachment | PdfSaveOptions.Attachments C# | embed text file in PDF | add OLE object to PDF | Aspose.Cells embed attachments | C# PDF attachment example | Aspose.Cells PdfSaveOptions | PDF with attached file
// Common Searches: how to attach a text file to a PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions embed attachments example C# | add OLE object as PDF attachment Aspose.Cells | C# create PDF with embedded text file Aspose | Aspose.Cells PDF attachment tutorial
// Developer Intent: Embed a text file into a PDF generated from an Aspose.Cells workbook using PdfSaveOptions.
// Use Cases: Include a log file as an attachment in a PDF report generated from spreadsheet data. | Embed a terms‑and‑conditions document within a financial statement PDF for compliance purposes. | Attach a specification sheet to a PDF invoice created from an Excel workbook.
// AI Prompts: Show me how to use PdfSaveOptions.Attachments to embed a text file in a PDF generated with Aspose.Cells (C#). | Provide C# code that creates a workbook, adds a text file as an OLE object, enables EmbedAttachments, and saves the PDF with the attachment. | Explain how to delete temporary files after embedding attachments in an Aspose.Cells PDF output.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfAttachmentDemo
{
    // Demonstrates how to create an Aspose.Cells workbook, generate a temporary text file, embed it as an OLE object, enable PdfSaveOptions.EmbedAttachments, and save the workbook as a PDF that contains the text file attachment. The example also shows how to clean up the temporary file after saving.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add a title.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("PDF with Text File Attachment");

                // Prepare a sample text file to be attached.
                string txtFilePath = "SampleAttachment.txt";
                File.WriteAllText(txtFilePath, "This is the content of the attached text file.");

                // Add the text file as an OLE object to the worksheet.
                // The OLE object will be embedded in the PDF when EmbedAttachments is true.
                int oleIndex = sheet.OleObjects.Add(10, 10, 200, 200, File.ReadAllBytes(txtFilePath));
                sheet.OleObjects[oleIndex].DisplayAsIcon = true; // Show as an icon

                // Create PDF save options and enable embedding of OLE attachments.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true
                };

                // Save the workbook as PDF.
                workbook.Save("WorkbookWithAttachment.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary text file if it exists.
                string txtFilePath = "SampleAttachment.txt";
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
