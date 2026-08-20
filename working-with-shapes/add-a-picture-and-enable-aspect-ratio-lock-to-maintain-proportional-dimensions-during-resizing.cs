// Title: C# – Insert a Picture into an Excel Worksheet and Preserve Its Aspect Ratio with Aspose.Cells
// Description: This example demonstrates how to create a new workbook, confirm the image file exists, add the picture to cell B2, enable the aspect‑ratio lock so the image scales proportionally, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells add image C# | lock picture aspect ratio .NET | insert picture Excel worksheet | maintain image proportions Aspose | C# Excel picture resizing | Aspose.Cells picture handling
// Common Searches: Aspose.Cells insert image keep aspect ratio | C# lock picture size when resizing Excel | how to add a logo to Excel with Aspose.Cells | prevent picture distortion in generated workbook | Aspose.Cells picture aspect ratio example
// Developer Intent: Add an image to a spreadsheet and ensure it scales without distortion.
// Use Cases: Embedding a company logo in automated reports while retaining its shape. | Displaying product thumbnails in a catalog sheet without stretching. | Applying a watermark that keeps its original proportions across different page sizes.
// AI Prompts: Write C# code that uses Aspose.Cells to place a PNG at cell B2 and enable proportional scaling. | Show how to toggle the IsAspectRatioLocked flag for an existing picture in a workbook. | Create robust error handling for missing image files when adding pictures with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example demonstrates how to create a new workbook, confirm the image file exists, add the picture to cell B2, enable the aspect‑ratio lock so the image scales proportionally, and save the result as an XLSX file using Aspose.Cells for .NET.
    public class AddPictureWithAspectRatioLock
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the image file
            string imagePath = "image.jpg";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            }

            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a picture to the worksheet
            int pictureIndex = worksheet.Pictures.Add(1, 1, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Lock the aspect ratio so the picture maintains its proportions when resized
            picture.IsAspectRatioLocked = true;

            // Save the workbook
            string outputPath = "PictureWithAspectRatioLock.xlsx";
            workbook.Save(outputPath);
        }
    }
}
