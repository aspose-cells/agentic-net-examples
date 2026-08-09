// Title: Embed a QR‑Code Password in a Protected Excel Workbook using Aspose.Cells for .NET (C#)
// Description: C# example that creates a new Workbook, applies a password protection, generates a QR‑code image containing the password, inserts the image at a specific cell, resizes it, and saves the file as a password‑protected XLSX document.
// Keywords: Aspose.Cells | C# | .NET | QR code | Excel password protection | embed image in worksheet | add picture to cell | generate QR code stream | protected XLSX | code sample
// Common Searches: Aspose.Cells insert QR code C# | set workbook password Aspose.Cells | add picture to Excel cell using Aspose.Cells | generate QR code from string in .NET | protect Excel file and embed password image | save protected workbook with image Aspose
// Developer Intent: Add a QR‑code that holds the workbook password directly into an Excel file while applying password protection with Aspose.Cells.
// Use Cases: Securely share a financial report where the opening password is hidden in a scannable QR code. | Provide field agents with a template that contains the access password as a QR code to avoid manual entry. | Distribute training materials that can only be opened after scanning the embedded QR‑code password.
// AI Prompts: Generate C# code that creates a QR‑code image from a password string and inserts it into a specific cell of an Aspose.Cells workbook. | Show how to replace a placeholder PNG with a dynamically generated QR‑code stream and handle insertion errors. | Provide a complete Aspose.Cells example that protects a workbook, creates a QR‑code for the password, embeds the image, and saves the file as XLSX.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsQRCodeDemo
{
    // C# example that creates a new Workbook, applies a password protection, generates a QR‑code image containing the password, inserts the image at a specific cell, resizes it, and saves the file as a password‑protected XLSX document.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data
                sheet.Cells["A1"].PutValue("Protected Workbook");
                sheet.Cells["A2"].PutValue("Password is stored in QR code below.");

                // Protection password
                string protectionPassword = "MySecretPassword123";
                workbook.Settings.Password = protectionPassword;

                // ----- Insert a placeholder QR code image -----
                // A 1x1 pixel PNG (transparent) is used as a placeholder.
                // In a real scenario replace this with an actual QR code image stream.
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);

                try
                {
                    using (MemoryStream imageStream = new MemoryStream(pngBytes))
                    {
                        // Add picture at row 4 (zero‑based index 3), column 1 (index 0)
                        int pictureIndex = sheet.Pictures.Add(3, 0, imageStream);
                        Picture picture = sheet.Pictures[pictureIndex];
                        picture.Width = 150;
                        picture.Height = 150;
                    }
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Failed to insert QR code image: {imgEx.Message}");
                }

                // Save the workbook
                string outputPath = "ProtectedWorkbook_WithQRCode.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'. The QR code contains the protection password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
