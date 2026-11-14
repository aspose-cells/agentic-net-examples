// Title: Create an Excel workbook and embed a QR code image from a Base64 string using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert image from base64 string into worksheet | how to add QR code bitmap to Excel using Aspose.Cells memory stream | convert base64 QR code to picture and place at A1 with Aspose.Cells | C# Aspose.Cells picture.Add example using MemoryStream | embed QR code image in Excel file without creating a temporary file Aspose
// Tags: base64 string to picture Aspose.Cells | picture.Add from MemoryStream Aspose.Cells | embed QR code bitmap in Excel worksheet | Aspose.Cells workbook image insertion | C# convert base64 to bitmap for Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample decodes a Base64‑encoded QR code, loads it into a MemoryStream, creates a new Workbook, adds the image as a picture to the first worksheet at cell A1 using the Pictures.Add method, and saves the workbook as QRCodeWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Base64 string representing the QR code image.
            // Replace the placeholder with the actual Base64 data.
            string base64Qr = "iVBORw0KGgoAAAANSUhEUgAA...";

            if (string.IsNullOrWhiteSpace(base64Qr))
                throw new ArgumentException("Base64 QR code data is missing.");

            // Convert the Base64 string to a byte array.
            byte[] imageBytes = Convert.FromBase64String(base64Qr);

            // Load the image bytes into a memory stream.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert the image into the worksheet as a picture.
                // The picture is placed at row 0, column 0 (top‑left corner).
                int pictureIndex = sheet.Pictures.Add(0, 0, ms);

                // Optional: adjust picture size if needed.
                Picture picture = sheet.Pictures[pictureIndex];
                // picture.Width = desiredWidth;
                // picture.Height = desiredHeight;

                // Define output file path.
                string outputPath = "QRCodeWorkbook.xlsx";

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
