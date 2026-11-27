// Title: Insert a JPEG picture into an Excel worksheet and crop its central area using Aspose.Cells for .NET (C#)
// Common Searches: C# code to add an image to Excel and crop the center with Aspose.Cells | how to trim equal margins from all sides of a picture in Aspose.Cells | Aspose.Cells picture cropping using UpperDeltaX and LowerDeltaY properties | set picture border and crop central portion in Excel via Aspose.Cells .NET | example of calculating percentage‑based crop values for Excel images in C#
// Tags: Aspose.Cells picture central cropping C# | UpperDeltaX LowerDeltaX picture trimming Aspose.Cells | insert JPEG into Excel worksheet Aspose.Cells | percentage based image crop Aspose.Cells .NET | add border to Excel picture Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureCropDemo
{
    // The sample loads a JPEG file, inserts it at cell B2 of a new workbook, computes 20% of the original width and height, applies those values to UpperDeltaX/Y and LowerDeltaX/Y to crop equal margins from all sides, adds a thin dark‑blue border, and saves the workbook as an Excel file with the centrally cropped picture.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source image file (ensure this file exists in the executable directory)
            const string imagePath = "sample.jpg";

            // Path for the output workbook
            const string outputPath = "CroppedPicture.xlsx";

            try
            {
                // Verify that the image file exists before proceeding
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Error: Image file \"{imagePath}\" not found.");
                    return;
                }

                // Create a new workbook and obtain the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add the picture to the worksheet.
                // Parameters: topRow, leftColumn, image file path.
                // Row and column indices are zero‑based. Here we place the picture starting at cell B2 (row 1, column 1).
                int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // ------------------------------------------------------------
                // Crop the picture so that only its central portion is visible.
                // The Picture class inherits from Shape, which provides the
                // UpperDeltaX/Y and LowerDeltaX/Y properties.
                // These properties define how many pixels are trimmed from each side.
                // To keep the central area, we trim the same amount from opposite sides.
                // ------------------------------------------------------------

                // Determine cropping amounts (example: trim 20% of width and height from each side)
                // First, obtain the original dimensions of the picture.
                int originalWidth = picture.OriginalWidth;
                int originalHeight = picture.OriginalHeight;

                // Calculate 20% of each dimension.
                int cropX = (int)(originalWidth * 0.20);
                int cropY = (int)(originalHeight * 0.20);

                // Apply cropping. UpperDeltaX/Y trim from the top‑left corner,
                // LowerDeltaX/Y trim from the bottom‑right corner.
                picture.UpperDeltaX = cropX;   // trim left side
                picture.LowerDeltaX = cropX;   // trim right side
                picture.UpperDeltaY = cropY;   // trim top side
                picture.LowerDeltaY = cropY;   // trim bottom side

                // Optional: set a thin border so the cropped area is clearly visible
                picture.BorderLineColor = System.Drawing.Color.DarkBlue;
                picture.BorderWeight = 1;

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Picture added and cropped successfully. Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
