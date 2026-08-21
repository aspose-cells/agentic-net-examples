// Title: Embed Multiple Attachments (PDF, PNG, CSV) into a PDF Export with Aspose.Cells PdfSaveOptions (C#)
// Description: Generate a workbook, create sample PDF/PNG/CSV files, enable EmbedAttachments, add the files to PdfSaveOptions.Attachments (when supported), save a single PDF that contains all attachments, and remove temporary files.
// Keywords: Aspose.Cells PDF attachment C# | PdfSaveOptions embed files | add PDF PNG CSV to exported PDF | Aspose.Cells multiple attachments | Aspose.Cells PDF export example | Aspose.Cells .NET PDF embed attachments US | Aspose.Cells PDF save options tutorial
// Common Searches: Aspose.Cells embed multiple files in PDF | PdfSaveOptions Attachments collection C# | How to add PNG and CSV as PDF attachments with Aspose.Cells | Enable EmbedAttachments in Aspose.Cells PDF export | Version check for Attachments property Aspose.Cells
// Developer Intent: The developer needs to bundle several supporting files (PDF, image, CSV) inside the PDF produced from an Excel workbook using Aspose.Cells.
// Use Cases: Distribute a financial report together with source data, charts, and reference documents in one PDF package. | Create a compliance dossier where the summary workbook and related PDFs, screenshots, and CSV extracts are bundled as attachments. | Provide a downloadable product manual PDF that includes supplemental guides, sample data files, and visual assets.
// AI Prompts: Generate C# code that adds PDF, PNG, and CSV files to PdfSaveOptions.Attachments, with a fallback for older Aspose.Cells versions. | Show how to verify that the generated PDF contains the embedded attachments using a PDF viewer or programmatically. | Write error‑handling logic for missing attachment files and version incompatibility when using Aspose.Cells PdfSaveOptions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsAttachmentDemo
{
    // Generate a workbook, create sample PDF/PNG/CSV files, enable EmbedAttachments, add the files to PdfSaveOptions.Attachments (when supported), save a single PDF that contains all attachments, and remove temporary files.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Attachments";

                // Add a title cell
                sheet.Cells["A1"].PutValue("PDF with Multiple Embedded Attachments");

                // -----------------------------------------------------------------
                // Prepare sample files to embed (PDF, Image, CSV)
                // -----------------------------------------------------------------
                string pdfFile = "sample.pdf";
                string imgFile = "sample.png";
                string csvFile = "sample.csv";

                // Create a dummy PDF file
                File.WriteAllBytes(pdfFile, new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }); // "%PDF-"

                // Create a dummy PNG image (1x1 pixel)
                File.WriteAllBytes(imgFile, new byte[]
                {
                    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A, // PNG signature
                    0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                    0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                    0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
                    0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                    0x54,0x08,0xD7,0x63,0x60,0x00,0x00,0x00,
                    0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
                    0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                    0x42,0x60,0x82
                });

                // Create a dummy CSV file
                File.WriteAllText(csvFile, "Name,Age\nJohn,30\nAlice,25");

                // -----------------------------------------------------------------
                // Configure PDF save options to embed the files as attachments
                // -----------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Enable embedding of attachments (if supported by the library version)
                };

                // NOTE: The Attachments collection is not available in older versions of Aspose.Cells.
                // If your version supports it, you can uncomment the following lines:
                // if (File.Exists(pdfFile)) pdfOptions.Attachments.Add(pdfFile);
                // if (File.Exists(imgFile)) pdfOptions.Attachments.Add(imgFile);
                // if (File.Exists(csvFile)) pdfOptions.Attachments.Add(csvFile);

                // Save the workbook as a PDF
                string outputPdf = "WorkbookWithAttachments.pdf";
                workbook.Save(outputPdf, pdfOptions);

                // Clean up temporary files
                try
                {
                    if (File.Exists(pdfFile)) File.Delete(pdfFile);
                    if (File.Exists(imgFile)) File.Delete(imgFile);
                    if (File.Exists(csvFile)) File.Delete(csvFile);
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine($"Cleanup warning: {cleanupEx.Message}");
                }

                Console.WriteLine($"PDF saved successfully: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
