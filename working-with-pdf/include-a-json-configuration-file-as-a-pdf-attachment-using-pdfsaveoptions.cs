// Title: Embed a JSON File as an Attachment in a PDF using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, generate a JSON configuration file, add it as an OLE object (icon) to a worksheet, enable PdfSaveOptions.EmbedAttachments, and save the workbook as a PDF that contains the JSON file as an embedded attachment.
// Keywords: Aspose.Cells PDF attachment | C# embed JSON in PDF | PdfSaveOptions EmbedAttachments | add OLE object worksheet | .NET generate PDF with attachment | Aspose.Cells embed file | export workbook to PDF with attachment
// Common Searches: Aspose.Cells embed JSON in PDF C# | PdfSaveOptions EmbedAttachments example | how to add OLE object to worksheet Aspose.Cells | save workbook as PDF with attached file | C# export Excel to PDF with attachment
// Developer Intent: Add a JSON configuration file to a PDF generated from an Aspose.Cells workbook by using OLE objects and the EmbedAttachments option.
// Use Cases: Distribute a PDF report that includes a downloadable JSON settings file for downstream automation. | Create an invoice PDF that carries a JSON payload with order details for ERP integration. | Provide a user guide PDF that embeds a JSON template for easy import into an application.
// AI Prompts: Generate C# code with Aspose.Cells to embed a JSON file as an OLE object and export the workbook to a PDF with EmbedAttachments enabled. | Explain the role of PdfSaveOptions.EmbedAttachments and the steps required to ensure the JSON file appears as an attachment in the resulting PDF. | List troubleshooting actions when the JSON attachment is missing from the exported PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfAttachmentDemo
{
    // Demonstrates how to create a workbook, generate a JSON configuration file, add it as an OLE object (icon) to a worksheet, enable PdfSaveOptions.EmbedAttachments, and save the workbook as a PDF that contains the JSON file as an embedded attachment.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a title to the worksheet (optional)
                worksheet.Cells["A1"].PutValue("PDF with JSON Attachment Example");

                // Prepare a JSON configuration file
                string jsonFilePath = "config.json";
                string jsonContent = @"{
    ""Setting1"": ""Value1"",
    ""Setting2"": 123,
    ""Enabled"": true,
    ""Items"": [""ItemA"", ""ItemB"", ""ItemC""]
}";
                File.WriteAllText(jsonFilePath, jsonContent);

                // Ensure the JSON file exists before embedding
                if (!File.Exists(jsonFilePath))
                    throw new FileNotFoundException("JSON configuration file was not created.", jsonFilePath);

                // Embed the JSON file as an OLE object (attachment) in the worksheet
                // Parameters: row, column, width, height, file bytes
                int oleIndex = worksheet.OleObjects.Add(5, 0, 200, 200, File.ReadAllBytes(jsonFilePath));
                // Display the attachment as an icon
                worksheet.OleObjects[oleIndex].DisplayAsIcon = true;
                // Note: IconCaption property is not available in the current Aspose.Cells version

                // Configure PDF save options to embed attachments
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true // Enable embedding of OLE attachments
                };

                // Save the workbook as a PDF file with the embedded JSON attachment
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
