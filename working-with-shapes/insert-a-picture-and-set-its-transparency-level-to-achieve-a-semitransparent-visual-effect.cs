// Title: Insert a Picture with Semi‑Transparent Effect in Excel using Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to add a PNG image to the first worksheet of a new Excel workbook, set its FormatPicture.Transparency to 0.5 for a 50 % semi‑transparent look, save the file, and verify the transparency after reloading. The code also creates a placeholder image if the source file is missing.
// Keywords: Aspose.Cells | C# | insert picture Excel | picture transparency | FormatPicture.Transparency | semi transparent image | Excel watermark | Aspose.Cells API | Excel workbook image | adjust picture opacity
// Common Searches: how to set picture transparency in Excel with Aspose.Cells C# | Aspose.Cells insert image with opacity | semi transparent picture in Excel using .NET | Aspose.Cells FormatPicture.Transparency example | C# code to add watermark image to Excel workbook
// Developer Intent: Add an image to a worksheet and apply a specific transparency level.
// Use Cases: Create a light watermark by inserting a logo at 50 % opacity behind data. | Overlay a semi‑transparent background on a chart to enhance visual contrast. | Generate reports where product photos need partial see‑through to keep text readable.
// AI Prompts: Show C# code that inserts a PNG into an Excel sheet and sets its transparency to 30 % using Aspose.Cells. | Write a method to open an existing workbook, change all picture opacities to 75 % and save it with Aspose.Cells for .NET. | Explain how to read and confirm the transparency value of a picture after loading a workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureTransparencyDemo
{
    // This example demonstrates how to add a PNG image to the first worksheet of a new Excel workbook, set its FormatPicture.Transparency to 0.5 for a 50 % semi‑transparent look, save the file, and verify the transparency after reloading. The code also creates a placeholder image if the source file is missing.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the image file
                string imagePath = "sampleImage.png";

                // Ensure the image exists; create a placeholder if it does not
                if (!File.Exists(imagePath))
                {
                    CreatePlaceholderImage(imagePath);
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the picture to the worksheet at row 2, column 2 (zero‑based indexes)
                int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Set picture transparency (0.0 = opaque, 1.0 = fully transparent)
                picture.FormatPicture.Transparency = 0.5; // 50% transparent

                // Save the workbook
                string outputPath = "PictureTransparencyDemo.xlsx";
                workbook.Save(outputPath);

                // Load the workbook again to verify the transparency setting
                if (File.Exists(outputPath))
                {
                    Workbook loadedWorkbook = new Workbook(outputPath);
                    Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
                    double loadedTransparency = loadedWorksheet.Pictures[0].FormatPicture.Transparency;
                    Console.WriteLine("Loaded picture transparency: " + loadedTransparency);
                }
                else
                {
                    Console.WriteLine("Failed to save the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        // Creates a simple 1x1 red PNG image as a placeholder
        private static void CreatePlaceholderImage(string path)
        {
            try
            {
                // PNG data for a 1x1 red pixel
                byte[] pngData = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
                File.WriteAllBytes(path, pngData);
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to create placeholder image: " + e.Message);
                throw;
            }
        }
    }
}
