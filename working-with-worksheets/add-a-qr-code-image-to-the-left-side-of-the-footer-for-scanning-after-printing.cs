// Title: Add a QR‑code image to the left footer of an Excel worksheet using Aspose.Cells for .NET
// Description: Shows how to read a PNG QR‑code from disk, embed it in the left section of a worksheet footer with SetFooterPicture, add the “&G” placeholder via SetFooter, and save the workbook, while handling a missing image file gracefully.
// Keywords: Aspose.Cells | .NET | C# | QR code | footer picture | SetFooterPicture | Excel printing | add image to footer | worksheet footer | load image from file | byte array
// Common Searches: Aspose.Cells add QR code to Excel footer | SetFooterPicture left section C# | Insert image in Excel footer using Aspose.Cells | How to print QR code in Excel footer .NET | Footer picture placeholder &G Aspose.Cells
// Developer Intent: Place a scannable QR‑code in the left side of a worksheet footer so it appears on printed pages.
// Use Cases: Printable reports that include a QR‑code linking to online documentation or dashboards. | Invoices where the footer QR‑code encodes payment details for quick mobile scanning. | Batch‑generated spreadsheets that embed a QR‑code for version tracking or source system identification. | Marketing sheets that provide a QR‑code for customers to access promotional content directly from the printed page.
// AI Prompts: Generate C# code that loads a PNG QR‑code and sets it as the left footer picture with Aspose.Cells. | Explain how to use SetFooterPicture and the “&G” placeholder to embed an image in an Excel footer, including error handling for missing files. | Show how to adjust the size or position of a QR‑code placed in the left footer using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Shows how to read a PNG QR‑code from disk, embed it in the left section of a worksheet footer with SetFooterPicture, add the “&G” placeholder via SetFooter, and save the workbook, while handling a missing image file gracefully.
    class AddQrCodeFooter
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                string qrPath = "qr.png";

                if (File.Exists(qrPath))
                {
                    // Load the QR code image into a byte array
                    byte[] qrImageData = File.ReadAllBytes(qrPath);

                    // Set the QR code image to the left section of the footer (section index 0)
                    worksheet.PageSetup.SetFooterPicture(0, qrImageData);

                    // Add the image placeholder script to the left footer so the image is displayed
                    worksheet.PageSetup.SetFooter(0, "&G");
                }
                else
                {
                    Console.WriteLine($"QR code image file '{qrPath}' not found. Footer will be created without image.");
                }

                // Save the workbook with the QR code footer
                string outputPath = "WorkbookWithQrFooter.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
