// Title: Move a picture shape to the back of an Excel worksheet using Aspose.Cells for .NET
// Description: Demonstrates how to add a PNG image to cell B2, send it behind all other worksheet objects with the ToFrontOrBack(-1) method, and save the workbook as PictureBackDemo.xlsx.
// Keywords: Aspose.Cells picture back | ToFrontOrBack negative | C# picture z‑order | Excel shape send to back | .NET insert image worksheet
// Common Searches: Aspose.Cells send picture to back C# | ToFrontOrBack(-1) example | move Excel shape behind cells using Aspose | place image behind other objects Aspose.Cells
// Developer Intent: Place an inserted picture behind every other object on the worksheet by adjusting its z‑order.
// Use Cases: Create a watermark that stays under cell data | Add a background illustration without covering labels | Arrange overlapping images where some need to be hidden behind others
// AI Prompts: Show C# code that inserts a PNG into a worksheet and moves it to the back with Aspose.Cells. | Explain the effect of positive and negative values in the ToFrontOrBack method. | Provide a step‑by‑step guide to change the z‑order of a picture shape in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a PNG image to cell B2, send it behind all other worksheet objects with the ToFrontOrBack(-1) method, and save the workbook as PictureBackDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path of the image to insert.
            string imagePath = "sample.png";

            // Create a simple placeholder PNG if the file does not exist.
            if (!File.Exists(imagePath))
            {
                // 1x1 pixel transparent PNG.
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lZL9WQAAAABJRU5ErkJggg==";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                File.WriteAllBytes(imagePath, pngBytes);
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add the picture to the worksheet at cell B2 (row 1, column 1).
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Send the picture to the back of the z-order.
            picture.ToFrontOrBack(-1);

            // Save the workbook.
            string outputPath = "PictureBackDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
