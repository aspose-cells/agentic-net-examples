// Title: Aspose.Cells Camera Shape – Capture Excel Range as High‑Resolution PNG in C#
// Description: Creates a workbook, populates cells A1:C3, captures the range with Pictures.Camera, positions the camera shape at row 5 column 1, sets ImageOrPrintOptions to 300 dpi PNG, and saves both the image and the workbook.
// Keywords: Aspose.Cells | Camera shape | C# export range to PNG | Excel range image | 300 DPI | ImageOrPrintOptions | .NET | high resolution Excel snapshot | Pictures.Camera method | save workbook as image
// Common Searches: Aspose.Cells C# capture range as PNG | Camera shape export Excel range with DPI | How to set image resolution in Aspose.Cells | Save worksheet range to high‑resolution image C# | Pictures.Camera method example
// Developer Intent: Generate an image of a specific worksheet range with custom resolution and format using the Camera shape.
// Use Cases: Create a printable thumbnail of a data table for reports. | Embed a crisp PNG snapshot of a chart or table in PDFs or web pages. | Automate batch export of multiple worksheet sections to images with consistent DPI.
// AI Prompts: Write C# code that captures range B2:D5 as a JPEG at 200 dpi using Aspose.Cells Camera shape. | Explain how to resize and reposition a Camera shape after capturing a range. | Provide a loop that exports a list of ranges to separate PNG files with varying resolutions.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsDemo
{
    // Creates a workbook, populates cells A1:C3, captures the range with Pictures.Camera, positions the camera shape at row 5 column 1, sets ImageOrPrintOptions to 300 dpi PNG, and saves both the image and the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data.
                worksheet.Cells["A1"].PutValue("Header1");
                worksheet.Cells["B1"].PutValue("Header2");
                worksheet.Cells["C1"].PutValue("Header3");
                worksheet.Cells["A2"].PutValue(10);
                worksheet.Cells["B2"].PutValue(20);
                worksheet.Cells["C2"].PutValue(30);
                worksheet.Cells["A3"].PutValue(40);
                worksheet.Cells["B3"].PutValue(50);
                worksheet.Cells["C3"].PutValue(60);

                // Capture picture of range A1:C3 and place it at row 5, column 1.
                PictureCollection pictures = worksheet.Pictures;
                int pictureIndex = pictures.Camera(5, 1, "A1:C3");
                Picture cameraShape = pictures[pictureIndex];

                // Define image rendering options.
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    ImageType = ImageType.Png
                };

                // Save the picture as an image file.
                string imagePath = "CameraCapture.png";
                cameraShape.ToImage(imagePath, imgOptions);
                Console.WriteLine($"Image saved to {imagePath}");

                // Save the workbook.
                string workbookPath = "CameraDemo.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved to {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
