// Title: C# – Embed a Text File as PDF Attachment with Aspose.Cells PdfSaveOptions
// Description: Creates a workbook, adds a temporary text file as an OLE object displayed as an icon, enables EmbedAttachments in PdfSaveOptions, and saves the workbook as a PDF that contains the text file as an embedded attachment.
// Keywords: Aspose.Cells PDF attachment | PdfSaveOptions EmbedAttachments | C# embed OLE object | Excel to PDF with attachment | Aspose.Cells embed text file
// Common Searches: embed text file in PDF using Aspose.Cells | PdfSaveOptions embed attachments .NET | add OLE object as icon in PDF with Aspose | save Excel workbook as PDF with attached file | Aspose.Cells PDF attachment example
// Developer Intent: Add a text document as an embedded attachment inside a PDF generated from an Excel workbook.
// Use Cases: Attach a terms‑and‑conditions file to a financial report PDF. | Include a README.txt with an invoice PDF for extra instructions. | Provide supplemental data as an embedded text file in a data‑analysis PDF export.
// AI Prompts: Generate C# code that embeds multiple CSV files as separate attachments in a PDF using Aspose.Cells. | Explain how to change the icon displayed for an embedded OLE object in the PDF. | List troubleshooting steps when the embedded attachment is missing from the saved PDF.

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds a temporary text file as an OLE object displayed as an icon, enables EmbedAttachments in PdfSaveOptions, and saves the workbook as a PDF that contains the text file as an embedded attachment.
class EmbedTextAsPdfAttachment
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("PDF with Embedded Text Attachment");

            // Create a temporary text file that will be embedded
            string textFilePath = "sample.txt";
            File.WriteAllText(textFilePath, "This is a sample text document to embed as an attachment.");

            // Ensure the file exists before reading its bytes
            if (!File.Exists(textFilePath))
                throw new FileNotFoundException("Temporary text file was not created.", textFilePath);

            // Add the text file as an OLE object (attachment) to the worksheet
            int oleIndex = worksheet.OleObjects.Add(10, 10, 200, 200, File.ReadAllBytes(textFilePath));
            // Display the OLE object as an icon; file format type is inferred automatically
            worksheet.OleObjects[oleIndex].DisplayAsIcon = true;

            // Configure PDF save options to embed OLE attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as a PDF file with the embedded text attachment
            string pdfPath = "PdfWithEmbeddedText.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully: {pdfPath}");

            // Delete the temporary text file
            if (File.Exists(textFilePath))
                File.Delete(textFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
