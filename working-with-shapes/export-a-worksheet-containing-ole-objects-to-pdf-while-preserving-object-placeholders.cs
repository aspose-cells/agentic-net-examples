// Title: Export Excel Worksheet with OLE Object Placeholders to PDF using Aspose.Cells for .NET
// Description: Demonstrates how to add an OLE object placeholder (using a 1×1 transparent PNG) to a worksheet, set the actual file bytes, configure DisplayDrawingObjects.Placeholders, disable attachment embedding with PdfSaveOptions, and save the sheet as a PDF while cleaning up temporary files.
// Keywords: Aspose.Cells | C# | .NET | export OLE objects to PDF | OLE placeholder | DisplayDrawingObjects.Placeholders | PdfSaveOptions EmbedAttachments false | Excel to PDF conversion | temporary file cleanup
// Common Searches: Aspose.Cells export OLE objects as placeholders PDF | C# add OLE object placeholder Excel Aspose.Cells | DisplayDrawingObjects.Placeholders PDF conversion | prevent embedding OLE data in PDF Aspose.Cells | how to clean up temporary files after Excel to PDF conversion
// Developer Intent: Create a PDF from an Excel worksheet that contains OLE objects, showing only their placeholder icons and omitting the embedded data.
// Use Cases: Generate lightweight PDF reports where embedded documents are represented by icons. | Export confidential spreadsheets while hiding actual OLE content for compliance. | Produce printable PDFs from templates that include OLE icons without increasing file size.
// AI Prompts: Write C# code with Aspose.Cells to add an OLE placeholder and export the worksheet to PDF without embedding the OLE data. | Explain the impact of DisplayDrawingObjects.Placeholders and PdfSaveOptions.EmbedAttachments on PDF output for OLE objects. | Provide a step‑by‑step guide for removing temporary files after converting a worksheet with OLE placeholders to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an OLE object placeholder (using a 1×1 transparent PNG) to a worksheet, set the actual file bytes, configure DisplayDrawingObjects.Placeholders, disable attachment embedding with PdfSaveOptions, and save the sheet as a PDF while cleaning up temporary files.
class ExportOleObjectsToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a temporary text file to be embedded as an OLE object
            string tempFilePath = "sample.txt";
            File.WriteAllText(tempFilePath, "This is sample content for the OLE object.");

            // Placeholder image (1x1 transparent PNG) required by Aspose.Cells when adding OLE objects
            byte[] placeholderImage = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

            // Add an OLE object placeholder using the generated image data
            int oleIndex = worksheet.OleObjects.Add(5, 2, 120, 120, placeholderImage);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Set the embedded object data (the actual file bytes) if the file exists
            if (File.Exists(tempFilePath))
            {
                oleObject.ObjectData = File.ReadAllBytes(tempFilePath);
            }

            // Optional: display as an icon with a label
            oleObject.DisplayAsIcon = true;
            oleObject.Label = "Sample Text File";

            // Configure the workbook to show placeholders instead of the actual OLE objects when rendering
            workbook.Settings.DisplayDrawingObjects = DisplayDrawingObjects.Placeholders;

            // Prepare PDF save options – ensure attachments are not embedded
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = false // placeholders will be kept
            };

            // Save the worksheet as a PDF file
            workbook.Save("WorksheetWithOlePlaceholders.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            // Clean up the temporary file if it exists
            string tempFilePath = "sample.txt";
            if (File.Exists(tempFilePath))
            {
                try
                {
                    File.Delete(tempFilePath);
                }
                catch (Exception delEx)
                {
                    Console.WriteLine("Failed to delete temporary file: " + delEx.Message);
                }
            }
        }
    }
}
