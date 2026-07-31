// Title: Embed a Macro‑Enabled Excel (XLSM) File as a PDF Attachment with Aspose.Cells for .NET (C#)
// Description: This example shows how to create or load an XLSM workbook, add it to another worksheet as an OLE object, enable PdfSaveOptions.EmbedAttachments, and save the result as a PDF that carries the macro file as an embedded attachment for reference.
// Keywords: Aspose.Cells embed OLE attachment PDF | C# embed XLSM in PDF | PdfSaveOptions EmbedAttachments | add Excel macro file to PDF | Aspose.Cells OLE object C# | macro‑enabled workbook PDF attachment | Aspose.Cells PDF export with attachment
// Common Searches: how to embed an xlsm file in a pdf using aspose.cells | c# add ole object to worksheet and export to pdf | aspose.cells PdfSaveOptions EmbedAttachments example | save workbook as pdf with embedded macro file | attach excel macro to pdf with aspose.cells
// Developer Intent: Generate a PDF that contains a macro‑enabled Excel workbook as an embedded attachment.
// Use Cases: Deliver audit‑ready PDF reports that include the original macro workbook for verification. | Provide documentation PDFs that bundle the source Excel macro for downstream editing or troubleshooting. | Automate batch conversion of workbooks to PDFs while preserving embedded macros for compliance archives.
// AI Prompts: Show C# code that adds a macro‑enabled Excel file as an OLE object to a worksheet and saves the workbook as a PDF with the attachment using Aspose.Cells. | Explain how to configure PdfSaveOptions.EmbedAttachments in Aspose.Cells to ensure OLE objects are embedded in the generated PDF. | Describe how to confirm that the macro file appears as an attachment in the resulting PDF with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAttachmentDemo
{
    // This example shows how to create or load an XLSM workbook, add it to another worksheet as an OLE object, enable PdfSaveOptions.EmbedAttachments, and save the result as a PDF that carries the macro file as an embedded attachment for reference.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the macro-enabled Excel file that will be attached to the PDF.
                string macroFilePath = "macro.xlsm";

                // Ensure the macro file exists; create a minimal one if it does not.
                if (!File.Exists(macroFilePath))
                {
                    Workbook macroWb = new Workbook();
                    // Save as macro-enabled workbook (Xlsm) so that it can be embedded later.
                    macroWb.Save(macroFilePath, SaveFormat.Xlsm);
                }

                // Create a simple workbook that will be converted to PDF.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("PDF with Embedded Macro Attachment");

                // Add the macro file as an OLE object (attachment) to the worksheet.
                // The OLE object is placed at row 10, column 10 with a size of 200x200 pixels.
                byte[] oleData = File.ReadAllBytes(macroFilePath);
                int oleIndex = sheet.OleObjects.Add(10, 10, 200, 200, oleData);
                OleObject ole = sheet.OleObjects[oleIndex];
                ole.FileFormatType = FileFormatType.Xlsm;   // Specify that the embedded file is an Excel macro workbook.
                ole.DisplayAsIcon = true;                  // Show the attachment as an icon.

                // Configure PDF save options to embed OLE attachments.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Enable embedding of the OLE object into the PDF.
                };

                // Save the workbook as a PDF file with the embedded macro attachment.
                string outputPdf = "WorkbookWithMacroAttachment.pdf";
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"PDF saved to '{outputPdf}' with the macro file embedded as an attachment.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
