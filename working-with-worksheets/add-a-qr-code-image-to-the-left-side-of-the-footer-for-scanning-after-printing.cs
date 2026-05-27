using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AddQrCodeToFooter
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the QR code image
            string qrImagePath = "qr_code.png";

            // Load QR code image if the file exists
            if (File.Exists(qrImagePath))
            {
                byte[] qrImageData = File.ReadAllBytes(qrImagePath);

                // Set the QR code image in the left footer (section index 0)
                // Returns a Picture object that can be further customized
                Picture footerPicture = worksheet.PageSetup.SetFooterPicture(0, qrImageData);

                // Insert the picture placeholder script into the left footer section
                // &G tells Excel to render the picture set for this section
                worksheet.PageSetup.SetFooter(0, "&G");

                // Optional: adjust picture properties (e.g., scaling) if desired
                // footerPicture.ScaleX = 0.5;
                // footerPicture.ScaleY = 0.5;
            }
            else
            {
                Console.WriteLine($"Warning: QR code image file '{qrImagePath}' not found. Footer will be saved without image.");
                // Optionally set a text placeholder in the footer
                worksheet.PageSetup.SetFooter(0, "QR code not available");
            }

            // Save the workbook
            string outputPath = "Workbook_With_QR_Footer.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}