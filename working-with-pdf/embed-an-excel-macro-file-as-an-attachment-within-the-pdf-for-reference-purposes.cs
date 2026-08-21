// Title: Embed an .xlsm Macro File as a PDF Attachment with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a macro‑enabled Excel file (.xlsm) as an OLE object, set its icon and label, enable attachment embedding via PdfSaveOptions, and save the workbook as a PDF that carries the macro file as an embedded attachment.
// Keywords: Aspose.Cells PDF attachment | embed .xlsm in PDF | C# OLE object PDF | PdfSaveOptions EmbedAttachments | macro file attachment Aspose | Aspose.Cells PDF export | .NET Excel to PDF with macro | global Aspose.Cells example
// Common Searches: Aspose.Cells embed macro file in PDF C# | PdfSaveOptions EmbedAttachments example | Add OLE object to worksheet and export to PDF | How to attach .xlsm to PDF using Aspose.Cells | C# export Excel with macro as PDF attachment
// Developer Intent: Create a PDF that contains a macro‑enabled Excel file as an embedded attachment using Aspose.Cells.
// Use Cases: Distribute a PDF report together with the original .xlsm macro for downstream analysis. | Provide documentation PDFs that include the supporting macro for audit trails. | Package a single PDF file that bundles both the rendered workbook and its executable macro for version‑controlled delivery.
// AI Prompts: Write C# code that embeds a .xlsm file as an OLE object and saves the workbook as a PDF with the macro attached using Aspose.Cells. | Show how to attach multiple macro files to a PDF with Aspose.Cells PdfSaveOptions. | Explain how to customize the icon and label of an embedded macro file when exporting to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a macro‑enabled Excel file (.xlsm) as an OLE object, set its icon and label, enable attachment embedding via PdfSaveOptions, and save the workbook as a PDF that carries the macro file as an embedded attachment.
class EmbedMacroInPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("PDF with Embedded Macro Attachment");

            // Path to the macro-enabled Excel file that will be embedded
            string macroFilePath = "macro.xlsm";

            // Create a placeholder macro file (in real use, provide an actual .xlsm file)
            File.WriteAllText(macroFilePath, "Placeholder content for macro file.");

            // Verify the macro file exists before embedding
            if (!File.Exists(macroFilePath))
                throw new FileNotFoundException("Macro file not found.", macroFilePath);

            // Prepare a simple PNG image (1x1 transparent pixel) to use as the OLE object icon
            // This avoids the need for System.Drawing which may not be available on all platforms
            byte[] iconBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");

            // Add an OLE object to embed the macro file
            // Parameters: topRow, leftColumn, height (pixels), width (pixels), imageData (icon)
            int oleIndex = worksheet.OleObjects.Add(5, 1, 200, 200, iconBytes);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the macro file data, display it as an icon with a label
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: File.ReadAllBytes(macroFilePath),
                sourceFileName: Path.GetFileName(macroFilePath),
                displayAsIcon: true,
                label: "Macro File"
            );

            // Specify the file format type so the correct icon can be shown (optional)
            oleObject.FileFormatType = FileFormatType.Xlsm;

            // Configure PDF save options to embed attachments (OLE objects) into the PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as PDF with the embedded macro attachment
            string pdfPath = "WorkbookWithMacroAttachment.pdf";
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            // Clean up the temporary macro file if it exists
            string macroFilePath = "macro.xlsm";
            if (File.Exists(macroFilePath))
            {
                try
                {
                    File.Delete(macroFilePath);
                }
                catch
                {
                    // Ignored – cleanup failure should not crash the program
                }
            }
        }
    }
}
