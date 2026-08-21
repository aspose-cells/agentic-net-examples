// Title: C# – Add a picture to an Excel worksheet and crop to the central area with Aspose.Cells for .NET
// Description: Creates a new Workbook, inserts a JPEG into cell B2, crops 25 % from each side (leaving the central 50 % of the image) using FormatPicture cropping properties, and saves the file as CroppedPicture.xlsx. Includes a file‑existence check and error handling.
// Keywords: Aspose.Cells | C# | add picture to Excel | crop image Excel | FormatPicture | LeftCrop | RightCrop | TopCrop | BottomCrop | central crop | worksheet picture | sample.jpg | CroppedPicture.xlsx
// Common Searches: Aspose.Cells insert picture and crop center .NET | C# crop picture in Excel using FormatPicture properties | How to display only the middle part of an image in Aspose.Cells | Aspose.Cells picture cropping example C# | Excel worksheet add image and trim margins with Aspose
// Developer Intent: Insert an image into a worksheet and apply cropping so that only the central portion of the picture is visible.
// Use Cases: Embed a company logo in a report while removing surrounding whitespace. | Generate a product catalog that shows cropped thumbnails focused on the main visual element. | Create a template that displays scanned documents with margins hidden by central cropping.
// AI Prompts: Generate C# code with Aspose.Cells to place a picture at cell D4 and crop 15 % from each side. | Explain the purpose of FormatPicture.LeftCrop, RightCrop, TopCrop, and BottomCrop and how to calculate pixel‑based cropping values. | Show an example that loads an image, adds it to a worksheet, and crops it to display only the middle 40 % of the picture.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // Creates a new Workbook, inserts a JPEG into cell B2, crops 25 % from each side (leaving the central 50 % of the image) using FormatPicture cropping properties, and saves the file as CroppedPicture.xlsx. Includes a file‑existence check and error handling.
    class AddCroppedPicture
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                string imagePath = "sample.jpg";

                // Ensure the image file exists before adding it
                if (File.Exists(imagePath))
                {
                    // Add picture at cell B2 (row 1, column 1)
                    int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Crop 25% from each side, leaving the central 50%
                    picture.FormatPicture.LeftCrop = 0.25;
                    picture.FormatPicture.RightCrop = 0.25;
                    picture.FormatPicture.TopCrop = 0.25;
                    picture.FormatPicture.BottomCrop = 0.25;
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook
                workbook.Save("CroppedPicture.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
