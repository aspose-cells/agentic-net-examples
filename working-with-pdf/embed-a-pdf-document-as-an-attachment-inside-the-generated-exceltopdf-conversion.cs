using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class EmbedPdfAttachmentDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Excel to PDF with embedded PDF attachment");

            // Path to the PDF file that will be embedded
            string pdfPath = "sample.pdf";

            // Ensure the PDF file exists (create a minimal placeholder if necessary)
            if (!File.Exists(pdfPath))
            {
                // Minimal PDF header bytes to make a valid PDF file
                File.WriteAllBytes(pdfPath, new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34, 0x0A });
            }

            // Read the PDF file into a byte array
            byte[] pdfBytes = File.ReadAllBytes(pdfPath);

            // Add an OLE object to the worksheet.
            // Pass null for imageData to use the default OLE icon.
            int oleIndex = sheet.OleObjects.Add(5, 0, 200, 200, null);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Embed the PDF data, display it as an icon, and set a label
            ole.SetEmbeddedObject(
                linkToFile: false,                 // Do not link to external file
                objectData: pdfBytes,              // Embedded PDF data
                sourceFileName: Path.GetFileName(pdfPath), // Original file name
                displayAsIcon: true,               // Show as an icon in the sheet
                label: "Embedded PDF");            // Icon label

            // Create PDF save options and enable embedding of OLE attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as a PDF; the embedded PDF will be attached to the output PDF
            string outputPath = "ExcelWithEmbeddedPdf.pdf";
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}