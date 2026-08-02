// Title: Attach a JSON Config File to a PDF using Aspose.Cells PdfSaveOptions (C#)
// Description: Demonstrates how to create a workbook, generate a temporary JSON configuration file, embed it as an OLE object displayed as an icon, enable attachment embedding with PdfSaveOptions, save the workbook as a PDF that contains the JSON file, and clean up the temporary file. Ideal for delivering reports that carry their original settings.
// Keywords: Aspose.Cells PDF attachment | PdfSaveOptions EmbedAttachments C# | embed JSON in PDF Aspose.Cells | OLE object Excel to PDF | C# generate PDF with attached file | Aspose.Cells export with attachment | JSON file as PDF attachment
// Common Searches: How to embed a JSON file in a PDF using Aspose.Cells C# | Aspose.Cells PdfSaveOptions EmbedAttachments example | Add OLE object to Excel worksheet and export to PDF | C# attach configuration file to generated PDF | Aspose.Cells save workbook as PDF with embedded files
// Developer Intent: Add a JSON configuration file as an embedded attachment in a PDF generated from an Excel workbook.
// Use Cases: Distribute a PDF report that includes the original JSON settings for offline reference. | Provide a downloadable PDF that carries a data‑schema JSON file for developers to import later. | Create a user guide PDF that automatically contains the required configuration JSON as an attached OLE object.
// AI Prompts: Show how to embed multiple files as OLE objects and ensure they appear in the PDF with Aspose.Cells. | Explain how to set a custom icon and caption for an OLE attachment when saving to PDF. | Provide robust error handling for missing or inaccessible files when embedding them as PDF attachments.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAttachmentDemo
{
    // Demonstrates how to create a workbook, generate a temporary JSON configuration file, embed it as an OLE object displayed as an icon, enable attachment embedding with PdfSaveOptions, save the workbook as a PDF that contains the JSON file, and clean up the temporary file. Ideal for delivering reports that carry their original settings.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add some sample data to the worksheet
                sheet.Cells["A1"].PutValue("Demo: JSON config attached to PDF");
                sheet.Cells["A2"].PutValue("See the attachment for configuration details.");

                // Prepare a JSON configuration file
                string jsonFilePath = "config.json";
                string jsonContent = @"{
    ""SettingA"": true,
    ""SettingB"": 42,
    ""SettingC"": ""Sample value""
}";
                File.WriteAllText(jsonFilePath, jsonContent);

                // Embed the JSON file as an OLE object (attachment) in the worksheet
                // Parameters: row, column, width, height, file bytes
                int oleIndex = sheet.OleObjects.Add(5, 0, 200, 200, File.ReadAllBytes(jsonFilePath));
                // Show the attachment as an icon
                sheet.OleObjects[oleIndex].DisplayAsIcon = true;
                // Note: IconCaption property is not available in current Aspose.Cells API; omitted.

                // Configure PDF save options to embed attachments
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Enable embedding of OLE attachments in the PDF
                };

                // Save the workbook as a PDF with the attachment embedded
                string pdfOutputPath = "WorkbookWithJsonAttachment.pdf";
                workbook.Save(pdfOutputPath, pdfOptions);

                // Clean up the temporary JSON file
                if (File.Exists(jsonFilePath))
                {
                    File.Delete(jsonFilePath);
                }

                Console.WriteLine($"PDF saved to '{pdfOutputPath}' with JSON attachment embedded.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
