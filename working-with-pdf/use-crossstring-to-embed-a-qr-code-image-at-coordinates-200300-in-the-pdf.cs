// Title: Embed a QR code PNG at (200, 300) points in a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a QR‑code PNG from a file, adds it to the first worksheet of an Aspose.Cells workbook via a MemoryStream, sets the picture's Left to 200 and Top to 300 points, and saves the workbook as a PDF. | Show how to position an Aspose.Cells.Drawing.Picture at exact coordinates before exporting the workbook to PDF, including handling of missing image files. | Provide a step‑by‑step example of creating a PDF with an overlaid QR code image using Aspose.Cells, demonstrating picture insertion, coordinate adjustment, and PDF saving.
// Common Searches: Aspose.Cells C# place image at specific X Y coordinates before PDF conversion | how to add QR code to PDF using Aspose.Cells workbook | set picture left and top properties in Aspose.Cells worksheet | embed PNG image into Excel sheet and export as PDF with Aspose.Cells | Aspose.Cells picture positioning units points C#
// Tags: Aspose.Cells picture positioning PDF | C# embed QR code Aspose.Cells | set picture coordinates Aspose.Cells | export worksheet to PDF with image overlay | load image bytes into Aspose.Cells picture | Aspose.Cells MemoryStream image insertion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# program that reads a QR‑code PNG, inserts it as a picture on the first worksheet, positions the picture at 200 pts left and 300 pts top, and saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the QR code image file
            string qrImagePath = @"C:\Path\To\qr_code.png";

            // Verify that the QR code image exists
            if (!File.Exists(qrImagePath))
            {
                Console.WriteLine($"QR code image not found: {qrImagePath}");
                return;
            }

            // Load the QR code image bytes
            byte[] qrImageBytes = File.ReadAllBytes(qrImagePath);

            // Create a new workbook (Excel file)
            Workbook workbook = new Workbook();

            // Get the first worksheet where the image will be placed
            Worksheet sheet = workbook.Worksheets[0];

            // Add the QR code image to the worksheet using a MemoryStream
            int pictureIndex;
            using (MemoryStream ms = new MemoryStream(qrImageBytes))
            {
                pictureIndex = sheet.Pictures.Add(0, 0, ms);
            }

            // Retrieve the inserted picture object to adjust its position
            Picture qrPicture = sheet.Pictures[pictureIndex];

            // Set the picture's position in points (Left = 200, Top = 300)
            qrPicture.Left = 200;
            qrPicture.Top = 300;

            // Ensure the output directory exists
            string outputPath = @"C:\Path\To\output.pdf";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
