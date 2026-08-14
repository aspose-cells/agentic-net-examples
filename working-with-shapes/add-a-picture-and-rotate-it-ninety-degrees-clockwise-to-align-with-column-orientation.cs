// Title: C# – Add a Picture to an Excel Worksheet and Rotate 90° Clockwise with Aspose.Cells
// Description: Shows how to create a workbook, verify an image file, insert the picture at a specific cell, set its RotationAngle to 90°, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# picture insertion | Excel image rotation | RotationAngle | Add picture to worksheet | Aspose.Cells example | rotate image 90 degrees | Excel shape manipulation | .NET Excel graphics | picture rotation code
// Common Searches: Aspose.Cells add image to Excel C# | rotate picture 90 degrees Aspose.Cells | set picture rotation angle C# Excel | insert picture into specific cell Aspose.Cells | how to rotate Excel shape programmatically
// Developer Intent: Insert an image into a worksheet and rotate it 90° clockwise.
// Use Cases: Align a company logo with column headers in a financial report. | Place a scanned signature on a form and rotate it for proper orientation. | Embed a diagram in a data sheet and rotate it to fit the column layout without manual adjustment. | Add a rotated watermark to a worksheet for branding purposes.
// AI Prompts: Generate C# code that adds a JPEG picture to cell C3 in an Aspose.Cells workbook and rotates it 90 degrees clockwise. | Explain step‑by‑step how to check for an image file, insert it into a worksheet, set RotationAngle, and save the workbook using Aspose.Cells for .NET. | Show how to rotate multiple pictures with different angles in a loop using Aspose.Cells C#. | What properties control picture rotation and positioning in Aspose.Cells, and how can they be combined?

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, verify an image file, insert the picture at a specific cell, set its RotationAngle to 90°, and save the file using Aspose.Cells for .NET.
class AddRotatedPicture
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file (replace with an actual file path)
            string imagePath = "image.jpg";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to the worksheet at the desired cell range
            // Here we place it starting at row 2, column 2 (zero‑based indices)
            int pictureIndex = sheet.Pictures.Add(2, 2, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Rotate the picture 90 degrees clockwise
            picture.RotationAngle = 90;

            // Save the workbook
            string outputPath = "RotatedPicture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
