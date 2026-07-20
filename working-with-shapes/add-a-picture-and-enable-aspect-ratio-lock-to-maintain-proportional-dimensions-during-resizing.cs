// Title: C# – Insert a picture into an Aspose.Cells worksheet and lock its aspect ratio
// Description: Creates a new Workbook, verifies the image file, inserts the picture into a defined cell range using a FileStream, enables IsAspectRatioLocked to keep proportions during resizing, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells add picture C# | lock aspect ratio Aspose.Cells | insert image Excel worksheet .NET | picture shape proportional scaling | FileStream picture Aspose.Cells
// Common Searches: Aspose.Cells insert image keep aspect ratio | C# picture shape lock aspect ratio example | How to add a picture to Excel with Aspose.Cells | Proportional image resizing Aspose.Cells C#
// Developer Intent: Add an image to a worksheet and ensure it retains its original proportions when the user resizes it.
// Use Cases: Embedding a company logo in a generated report without distortion. | Adding product photos to a catalog where column width changes must not stretch the images. | Creating a template that lets end‑users drag picture corners while automatically preserving aspect ratio.
// AI Prompts: Provide C# code that inserts a picture into an Aspose.Cells worksheet and locks its aspect ratio. | Show how to check for an image file before adding it as a shape in an Aspose.Cells workbook. | Demonstrate inserting a picture with a FileStream into specific rows and columns and enabling proportional resizing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new Workbook, verifies the image file, inserts the picture into a defined cell range using a FileStream, enables IsAspectRatioLocked to keep proportions during resizing, and saves the file as an Excel workbook.
class AddPictureWithAspectRatioLock
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            string imagePath = "image.jpg";

            // Verify that the image file exists before attempting to load it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }
            else
            {
                // Open the image file as a stream and add it to the worksheet
                using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                {
                    // Parameters: topRow, leftColumn, bottomRow, rightColumn, image stream
                    Picture picture = worksheet.Shapes.AddPicture(2, 2, 10, 10, fs);
                    // Lock aspect ratio so the picture maintains its proportions when resized
                    picture.IsAspectRatioLocked = true;
                }
            }

            // Save the workbook with the picture whose aspect ratio is locked
            string outputPath = "OutputWithLockedAspectRatio.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
