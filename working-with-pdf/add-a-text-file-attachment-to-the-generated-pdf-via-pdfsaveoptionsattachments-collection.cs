// Title: Add a Text File Attachment to a PDF Using Aspose.Cells C#
// Description: Demonstrates how to create a workbook, generate a temporary text file, embed it as an OLE object, enable attachment embedding with PdfSaveOptions, and save the workbook as a PDF that contains the text file as an attachment. The sample also cleans up the temporary file after saving.
// Keywords: Aspose.Cells PDF attachment C# | PdfSaveOptions.Attachments collection | embed text file in PDF Aspose.Cells | C# add OLE object to PDF | Aspose.Cells save workbook as PDF with attachment | PDF embed file Aspose.Cells example
// Common Searches: how to embed a text file in a PDF with Aspose.Cells C# | Aspose.Cells PdfSaveOptions Attachments example | add OLE object as PDF attachment using Aspose.Cells | C# generate PDF with embedded files Aspose.Cells | Aspose.Cells embed attachment in exported PDF
// Developer Intent: Embed a plain‑text file inside the PDF produced from an Aspose.Cells workbook.
// Use Cases: Attach a log or audit file to a PDF report generated from spreadsheet data. | Include terms‑and‑conditions or policy documents with a PDF invoice created from Excel. | Provide supplemental CSV or TXT data as an embedded file within a PDF summary.
// AI Prompts: Write C# code that adds multiple files to PdfSaveOptions.Attachments instead of using OLE objects. | Explain when to use PdfSaveOptions.EmbedAttachments versus the Attachments collection in Aspose.Cells. | Show how to customize the icon displayed for an OLE object that becomes a PDF attachment.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsAttachmentDemo
{
    // Demonstrates how to create a workbook, generate a temporary text file, embed it as an OLE object, enable attachment embedding with PdfSaveOptions, and save the workbook as a PDF that contains the text file as an attachment. The sample also cleans up the temporary file after saving.
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

                // Prepare a simple text file to be attached
                string txtFilePath = "SampleAttachment.txt";
                File.WriteAllText(txtFilePath, "This is the content of the attached text file.");

                // Ensure the text file exists before adding it as an OLE object
                if (!File.Exists(txtFilePath))
                    throw new FileNotFoundException("Attachment file not found.", txtFilePath);

                // Add the text file as an OLE object to the worksheet
                // Parameters: upper left row, column, height, width, file bytes
                int oleIndex = sheet.OleObjects.Add(5, 0, 200, 200, File.ReadAllBytes(txtFilePath));

                // Display the OLE object as an icon
                sheet.OleObjects[oleIndex].DisplayAsIcon = true;

                // Configure PDF save options to embed OLE attachments
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Enables embedding of OLE objects as PDF attachments
                };

                // Save the workbook as PDF with the attachment embedded
                string outputPdf = "WorkbookWithAttachment.pdf";
                workbook.Save(outputPdf, pdfOptions);

                // Clean up the temporary text file
                File.Delete(txtFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
