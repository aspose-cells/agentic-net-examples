// Title: C# – Add and Center‑Crop a Picture in Excel with Aspose.Cells
// Description: Demonstrates how to create a workbook, insert an image into cell A1, apply 25 % left/right/top/bottom cropping to keep the central 50 % of the picture, and save the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells picture crop | C# add image Excel | Excel image cropping .NET | center crop Aspose.Cells | FormatPicture LeftCrop RightCrop | Aspose.Cells picture example
// Common Searches: how to crop a picture in Aspose.Cells | center crop image Excel C# | Aspose.Cells add picture to worksheet | C# crop picture percentages Aspose.Cells | remove borders from Excel image programmatically
// Developer Intent: Insert an image into a worksheet and trim it so only the central portion remains visible.
// Use Cases: Generate reports with logos that need only the central emblem displayed. | Create product catalogs where each photo is automatically centered and cropped to fit a cell. | Build dashboards that focus on the main content of pictures without manual editing.
// AI Prompts: Write C# code using Aspose.Cells to add a picture from a file path and crop 30 % from each side. | Explain the effect of LeftCrop, RightCrop, TopCrop, and BottomCrop on an Excel image placed with Aspose.Cells. | Provide a reusable method that accepts custom cropping percentages and applies them to any picture in a workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCropExample
{
    // Demonstrates how to create a workbook, insert an image into cell A1, apply 25 % left/right/top/bottom cropping to keep the central 50 % of the picture, and save the file as an XLSX document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the image file (replace with your actual image path)
                string imagePath = "sample.jpg";

                // Verify that the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}");
                    return;
                }

                // Add the picture to the worksheet at cell A1 (row 0, column 0)
                int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Crop the picture to show only the central portion.
                // Setting each side to 0.25 crops 25% from that side,
                // leaving the middle 50% of the image visible.
                picture.FormatPicture.LeftCrop = 0.25;
                picture.FormatPicture.RightCrop = 0.25;
                picture.FormatPicture.TopCrop = 0.25;
                picture.FormatPicture.BottomCrop = 0.25;

                // Save the workbook with the cropped picture
                string outputPath = "CroppedPictureDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
