using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class EmbedMacroInPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF with Embedded Macro");

        // Create a temporary macro-enabled workbook to embed
        string macroFilePath = "sample_macro.xlsm";
        Workbook macroWorkbook = new Workbook();
        macroWorkbook.Worksheets[0].Cells["A1"].PutValue("Macro content");
        macroWorkbook.Save(macroFilePath, SaveFormat.Xlsm);

        // Read macro file bytes
        byte[] macroData = File.ReadAllBytes(macroFilePath);

        // Add the macro as an embedded OLE object
        int oleIndex = worksheet.OleObjects.Add(5, 1, 200, 200, macroData);
        OleObject oleObject = worksheet.OleObjects[oleIndex];
        oleObject.FileFormatType = FileFormatType.Xlsm;
        oleObject.DisplayAsIcon = true;
        oleObject.Label = "Macro File";

        // Configure PDF save options to embed attachments
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.EmbedAttachments = true;

        // Save the workbook as PDF with the embedded macro attachment
        workbook.Save("WorkbookWithMacroAttachment.pdf", pdfOptions);

        // Clean up temporary macro file
        File.Delete(macroFilePath);
    }
}