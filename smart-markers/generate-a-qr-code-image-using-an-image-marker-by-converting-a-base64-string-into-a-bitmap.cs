// Title: Insert a QR code image from a Base64 string into an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that decodes a Base64‑encoded QR code, creates a picture via MemoryStream, inserts it into the first worksheet of an Aspose.Cells workbook, and saves the file as XLSX. | Show how to add an image to an Aspose.Cells worksheet using a byte array derived from a Base64 string and assign a custom name to the picture. | Provide a complete example that loads a Base64 QR‑code image, embeds it as a picture in an Excel sheet, and writes the workbook to the current directory.
// Common Searches: how to embed a base64 QR code image into an Excel file using Aspose.Cells C# | Aspose.Cells add picture from memory stream example | C# convert base64 string to bitmap and insert into worksheet | save workbook with QR code image using Aspose.Cells .NET
// Tags: base64 string to picture Aspose.Cells | insert image into worksheet using memory stream | embed QR code bitmap in XLSX C# | assign picture name Aspose.Cells | save workbook with image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates decoding a Base64‑encoded QR‑code, creating a picture via MemoryStream, adding it to the first worksheet of a new Aspose.Cells workbook, naming the picture, and saving the workbook as an XLSX file in the current directory.
class GenerateQrCodeImage
{
    static void Main()
    {
        try
        {
            // Base64 string of the QR code image (replace with your actual string)
            string base64QrCode = "iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAYAAAB..."; // truncated for brevity

            // Convert Base64 string to a byte array
            byte[] imageBytes = Convert.FromBase64String(base64QrCode);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the QR code image to the worksheet using a memory stream
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                int pictureIndex = sheet.Pictures.Add(0, 0, ms);
                Picture picture = sheet.Pictures[pictureIndex];
                picture.Name = "QrCodeImage";
            }

            // Define output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "QrCodeWorkbook.xlsx");

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
