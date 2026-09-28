// Title: Insert a QR‑code PNG decoded from a Base64 string into a specific cell of an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Decode the Base64‑encoded PNG, create a MemoryStream, and add the image to cell B3 of the first worksheet using Workbook.Worksheets[0].Pictures.Add. | Set the inserted picture's Width and Height properties to 150 pixels each to control its size. | Save the workbook to a file such as QrCodeWorkbook.xlsx and confirm the file is generated without errors.
// Common Searches: how to decode a base64 QR code image and place it in an Excel cell using Aspose.Cells C# | Aspose.Cells add picture from memory stream to specific cell | set picture size in Aspose.Cells workbook C# | embed PNG image from base64 string into Excel worksheet with Aspose.Cells | convert base64 string to bitmap for Excel using Aspose.Cells .NET
// Tags: add base64 decoded image to Aspose.Cells worksheet | embed QR code PNG in Excel using Aspose.Cells | insert picture from MemoryStream Aspose.Cells | resize picture dimensions Aspose.Cells C# | convert base64 to bitmap for Excel workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example decodes a Base64‑encoded PNG that contains a QR code, creates a MemoryStream, and inserts the image into cell B3 of the first worksheet using Aspose.Cells. The picture is resized to 150 × 150 pixels, and the workbook is saved as QrCodeWorkbook.xlsx.
class GenerateQrCodeWorkbook
{
    static void Main()
    {
        try
        {
            // Base64 string that represents the QR code image (PNG format in this example)
            string base64QrImage = "iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAYAAAB..."; // TODO: replace with actual Base64 data

            if (string.IsNullOrWhiteSpace(base64QrImage))
                throw new InvalidOperationException("Base64 QR image data is missing.");

            // Convert Base64 string to a byte array
            byte[] imageBytes = Convert.FromBase64String(base64QrImage);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add the QR code image to the worksheet directly from the byte stream
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                int pictureIndex = sheet.Pictures.Add(2, 1, ms); // places image at cell B3

                // Optionally adjust picture size
                Picture picture = sheet.Pictures[pictureIndex];
                picture.Width = picture.Height = 150; // example size, adjust as needed
            }

            // Define output file path
            string outputPath = "QrCodeWorkbook.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
