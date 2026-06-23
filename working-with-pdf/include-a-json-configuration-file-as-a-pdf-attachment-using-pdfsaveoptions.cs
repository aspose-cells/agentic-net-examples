using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF with JSON attachment");

        // Create a sample JSON configuration file
        string jsonFile = "config.json";
        File.WriteAllText(jsonFile, "{ \"setting\": true, \"value\": 123 }");

        // Add the JSON file as an OLE object (attachment) to the worksheet
        int oleIndex = worksheet.OleObjects.Add(5, 5, 150, 150, File.ReadAllBytes(jsonFile));
        worksheet.OleObjects[oleIndex].FileFormatType = FileFormatType.Unknown; // generic file type
        worksheet.OleObjects[oleIndex].DisplayAsIcon = true; // show as an icon

        // Configure PDF save options to embed attachments
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.EmbedAttachments = true;

        // Save the workbook as a PDF with the embedded JSON attachment
        workbook.Save("WorkbookWithJsonAttachment.pdf", pdfSaveOptions);

        // Clean up the temporary JSON file
        File.Delete(jsonFile);
    }
}