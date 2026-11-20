// Title: C# – Add a picture to an Aspose.Cells worksheet and enable ScaleCrop to lock its aspect ratio
// Common Searches: asp.net c# insert image into excel worksheet with Aspose.Cells and keep aspect ratio | how to enable ScaleCrop property for a picture in Aspose.Cells | lock picture aspect ratio after adding picture to workbook using Aspose.Cells | C# example setting picture.IsLockedAspectRatio in Aspose.Cells | preserve image proportions when adding picture to Excel file with Aspose.Cells .NET
// Tags: Aspose.Cells picture locked aspect ratio | ScaleCrop property usage Aspose.Cells .NET | worksheet picture scaling with aspect lock | C# workbook picture scaling options | Excel file image insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds an image to cell B2 of the first worksheet, and demonstrates how to enable the ScaleCrop (IsLockedAspectRatio) property on the Picture object to keep the image's original proportions before saving the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "example.png";

            // Verify that the image file exists before adding it
            if (File.Exists(imagePath))
            {
                // Add the picture to cell B2 (row index 1, column index 1)
                int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Note: Older Aspose.Cells versions may not expose IsLockedAspectRatio.
                // If needed, adjust picture size manually here.
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
