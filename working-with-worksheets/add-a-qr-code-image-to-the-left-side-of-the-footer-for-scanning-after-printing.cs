// Title: Add a QR Code to the Left Footer of an Excel Workbook using Aspose.Cells for .NET (C#)
// Description: This example shows how to load a QR‑code PNG, place it in the left section of a worksheet footer with `PageSetup.SetFooterPicture(0, …)`, insert the `&G` placeholder via `SetFooter(0, "&G")`, and save the workbook as `Workbook_With_QR_Footer.xlsx`.
// Keywords: Aspose.Cells QR code footer | SetFooterPicture C# | Excel left footer image | Add image to Excel footer .NET | PageSetup footer picture example | C# insert QR code in Excel | Aspose.Cells footer customization
// Common Searches: Aspose.Cells add QR code to left footer | C# set footer picture in Excel workbook | How to place an image in Excel footer using Aspose.Cells | Insert QR code in Excel footer programmatically | SetFooterPicture left section Aspose.Cells
// Developer Intent: Place a QR‑code image in the left part of a worksheet footer.
// Use Cases: Print‑ready reports that include a scannable QR code for quick URL access. | Invoices with a QR code in the footer encoding payment information or a payment‑portal link. | Shipping documents where the footer QR code points to real‑time tracking data.
// AI Prompts: Generate C# code that loads a PNG and sets it as the left footer picture with Aspose.Cells. | Show how to add different QR‑code images to both left and right footers of a worksheet. | Explain how to adjust size, scaling, and alignment of a footer picture using SetFooterPicture.

using System;
using System.IO;
using Aspose.Cells;

// This example shows how to load a QR‑code PNG, place it in the left section of a worksheet footer with `PageSetup.SetFooterPicture(0, …)`, insert the `&G` placeholder via `SetFooter(0, "&G")`, and save the workbook as `Workbook_With_QR_Footer.xlsx`.
class AddQrCodeFooter
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

            // Load the QR code image if the file exists
            if (File.Exists(qrImagePath))
            {
                byte[] qrImageData = File.ReadAllBytes(qrImagePath);

                // Set the QR code image in the left section of the footer (section index 0)
                worksheet.PageSetup.SetFooterPicture(0, qrImageData);

                // Add the image placeholder script to the left footer so the image is displayed
                worksheet.PageSetup.SetFooter(0, "&G");
            }
            else
            {
                Console.WriteLine($"QR code image file not found: {qrImagePath}. Footer image will be omitted.");
            }

            // Save the workbook with the QR code footer
            workbook.Save("Workbook_With_QR_Footer.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
