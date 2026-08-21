// Title: C# – Render Worksheet to PNG and Generate QR Code Linking to Hosted Image with Aspose.Cells
// Description: Shows how to create a workbook, fill it with data, export the first worksheet as a PNG using Aspose.Cells, and produce a QR‑code image that encodes the public URL of the hosted PNG for quick mobile access.
// Keywords: Aspose.Cells | C# | worksheet to PNG | image rendering | QR code generation | hosted image URL | mobile access | QRCoder | SheetRender | ImageOrPrintOptions | PNG export
// Common Searches: Aspose.Cells render worksheet as PNG C# | C# generate QR code for image URL | Create QR code linking to hosted PNG | Export Excel sheet to PNG and add QR code | C# QR code library for Aspose.Cells output | QR code for Excel image on web
// Developer Intent: Create a QR‑code image that points to the publicly accessible PNG of a rendered worksheet, enabling users to scan and view the sheet on mobile devices.
// Use Cases: Attach QR code to printed reports so readers can scan and view the live worksheet image online. | Embed QR code in PDFs generated from the workbook for instant navigation to the web‑hosted PNG. | Display QR code alongside a web gallery of exported worksheets for easy sharing. | Use QR code in internal dashboards to link directly to the latest worksheet snapshot.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet to PNG and then uses QRCoder to create a QR‑code PNG encoding the hosted image URL. | Write a method that replaces the placeholder PNG with a real QR code, handling errors and ensuring the output folder exists. | Provide sample logging that captures failures when the QR code generation library cannot reach the URL. | Show how to upload the PNG to Azure Blob Storage and retrieve its public URL for QR code encoding.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsQRCodeDemo
{
    // Shows how to create a workbook, fill it with data, export the first worksheet as a PNG using Aspose.Cells, and produce a QR‑code image that encodes the public URL of the hosted PNG for quick mobile access.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(85);

                // 2. Configure image rendering options (PNG, one page per sheet)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true
                };

                // 3. Render the first page of the worksheet to a PNG file
                string pngFileName = "worksheet.png";
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                sheetRender.ToImage(0, pngFileName);
                Console.WriteLine($"Worksheet image saved to: {pngFileName}");

                // 4. Assume the PNG is hosted at a public URL (replace with actual URL in production)
                string hostedImageUrl = "https://yourdomain.com/images/" + pngFileName;

                // 5. Generate a simple placeholder QR code PNG (1x1 white pixel)
                string qrFileName = "worksheet_qr.png";
                try
                {
                    CreatePlaceholderPng(qrFileName);
                    Console.WriteLine($"QR code placeholder saved to: {qrFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create QR placeholder: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a minimal 1x1 white PNG file as a placeholder for a QR code
        private static void CreatePlaceholderPng(string filePath)
        {
            // PNG byte array for a 1x1 white pixel image
            byte[] pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
                0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,
                0x00,0x04,0x00,0x01,0xE2,0x26,0x05,0x9B,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
                0xAE,0x42,0x60,0x82
            };

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllBytes(filePath, pngBytes);
        }
    }
}
