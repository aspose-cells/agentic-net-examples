// Title: Embed Multiple Files as PDF Attachments with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds two OLE objects (a text file and a DOCX file), enables PdfSaveOptions.EmbedAttachments, and saves the workbook as a PDF containing the external files as embedded attachments. Temporary files are removed after saving.
// Keywords: Aspose.Cells | PdfSaveOptions | EmbedAttachments | C# | multiple PDF attachments | OLE object | Excel to PDF | embed txt file | embed docx file | Aspose.Cells for .NET | PDF attachment | export workbook with attachments
// Common Searches: Aspose.Cells embed multiple attachments in PDF | PdfSaveOptions EmbedAttachments C# example | How to add OLE objects and save as PDF with attachments | Save Excel as PDF with embedded files using Aspose | C# code to embed txt and docx in PDF via Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook that bundles several external files as embedded attachments using Aspose.Cells for .NET.
// Use Cases: Create a single PDF report that includes supporting documents (e.g., a summary text file and a contract Word file) for easy distribution. | Automate invoice PDFs that carry attached terms‑and‑conditions or warranty documents alongside the invoice data. | Produce compliance packages where logs, policies, and reference files are packaged as embedded attachments within one PDF.
// AI Prompts: Show how to embed additional file types such as images or PDFs as attachments using Aspose.Cells PdfSaveOptions. | Refactor the sample to accept a list of file paths and embed them in a loop instead of hard‑coding each file. | Explain how to extract the embedded attachments from the generated PDF with Aspose.PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

namespace AsposeCellsPdfAttachmentDemo
{
    // C# example that creates a workbook, adds two OLE objects (a text file and a DOCX file), enables PdfSaveOptions.EmbedAttachments, and saves the workbook as a PDF containing the external files as embedded attachments. Temporary files are removed after saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("PDF with Multiple Embedded Attachments");

            // Prepare sample files to embed
            string txtFile = "SampleText.txt";
            string docxFile = "SampleDoc.docx";

            try
            {
                // Create a simple text file
                File.WriteAllText(txtFile, "This is a sample text file to be embedded.");

                // Create a simple docx file (placeholder bytes)
                File.WriteAllBytes(docxFile, new byte[] { 80, 75, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });

                // Add first OLE object (text file) to the worksheet
                if (File.Exists(txtFile))
                {
                    int oleIndex1 = worksheet.OleObjects.Add(5, 5, 200, 200, File.ReadAllBytes(txtFile));
                    OleObject oleObject1 = worksheet.OleObjects[oleIndex1];
                    oleObject1.DisplayAsIcon = true;
                    oleObject1.Label = Path.GetFileName(txtFile);
                }

                // Add second OLE object (docx file) to the worksheet
                if (File.Exists(docxFile))
                {
                    int oleIndex2 = worksheet.OleObjects.Add(15, 5, 200, 200, File.ReadAllBytes(docxFile));
                    OleObject oleObject2 = worksheet.OleObjects[oleIndex2];
                    oleObject2.DisplayAsIcon = true;
                    oleObject2.Label = Path.GetFileName(docxFile);
                }

                // Configure PDF save options to embed attachments
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true
                };

                // Save the workbook as PDF with embedded attachments
                string outputPdf = "PdfWithMultipleEmbeddedAttachments.pdf";
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"PDF saved successfully: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files
                if (File.Exists(txtFile)) File.Delete(txtFile);
                if (File.Exists(docxFile)) File.Delete(docxFile);
            }
        }
    }
}
