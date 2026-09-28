// Title: Generate a password‑protected Excel workbook and embed a QR‑code image at cell B2 using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new .xlsx file, protect its structure and windows with a password, insert a QR‑code PNG into cell B2, set the image size to 150 × 150 pixels, and save the workbook using Aspose.Cells in C#. | Load a QR‑code image file, add it to a protected worksheet at a specific cell, adjust the picture dimensions, and export the password‑protected workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# protect workbook structure and windows with password and add QR code image to cell B2 | How to insert a PNG picture into a specific cell after protecting an Excel file using Aspose.Cells | Set picture width and height in an Aspose.Cells worksheet C# | Save a password‑protected Excel file with an embedded QR code using Aspose.Cells .NET
// Tags: Aspose.Cells protect workbook with password C# | Aspose.Cells insert QR code image into worksheet | Aspose.Cells add picture to cell B2 | Aspose.Cells set picture size 150x150 | Aspose.Cells save protected workbook as xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, applies password protection to its structure and windows, attempts to load a QR‑code PNG from "qr.png" and places it at cell B2 with a size of 150 × 150 pixels (or writes a placeholder if missing), and saves the result as "ProtectedWorkbookWithQR.xlsx" using Aspose.Cells for .NET.
class GenerateQrProtectedWorkbook
{
    static void Main()
    {
        try
        {
            // Define the protection password
            string protectionPassword = "SecurePass123";

            // Create a new workbook and protect it (protect structure and windows)
            Workbook workbook = new Workbook();
            workbook.Protect(ProtectionType.All, protectionPassword);
            Worksheet sheet = workbook.Worksheets[0];

            // Optional: path to a pre‑generated QR code image
            string qrImagePath = "qr.png";

            if (File.Exists(qrImagePath))
            {
                // Add the QR code image to the worksheet at cell B2 (row 1, column 1)
                int pictureIndex = sheet.Pictures.Add(1, 1, qrImagePath);
                Picture qrPicture = sheet.Pictures[pictureIndex];
                qrPicture.Width = 150;
                qrPicture.Height = 150;
            }
            else
            {
                // Indicate that the QR image is missing
                sheet.Cells["B2"].PutValue("QR image not found.");
            }

            // Save the workbook to a file
            string outputPath = "ProtectedWorkbookWithQR.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
