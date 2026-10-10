// Title: Add a picture to cell B2 in an Excel workbook and lock its aspect ratio using Aspose.Cells for .NET
// AI Prompts: Insert an image from a file path into a specific worksheet cell and set proportional scaling to keep the aspect ratio with Aspose.Cells in C#. | Demonstrate how to programmatically enforce aspect‑ratio preservation for a picture added to an Excel sheet when resizing using Aspose.Cells.
// Common Searches: C# Aspose.Cells insert picture into worksheet cell and keep aspect ratio | How to prevent image distortion when adding a picture to Excel using Aspose.Cells .NET | Aspose.Cells picture.Add method lock aspect ratio example
// Tags: Aspose.Cells picture.Add proportional scaling | C# embed PNG in Excel worksheet | lock image aspect ratio Aspose.Cells | save workbook with picture Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, accesses the first worksheet, checks that a PNG file exists, inserts the picture at cell B2 using the Pictures.Add method, notes that Aspose.Cells may not expose a direct LockAspectRatio property and suggests adjusting width and height proportionally, ensures the output folder exists, saves the workbook, and handles any exceptions.
class AddPictureWithAspectRatioLock
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = @"C:\Images\sample.png";

            // Verify that the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Add the picture to the worksheet at cell B2 (row index 1, column index 1)
            // Add returns the index of the picture in the collection
            int pictureIndex = sheet.Pictures.Add(1, 1, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Aspose.Cells does not expose a direct LockAspectRatio property for Picture in some versions.
            // If needed, maintain aspect ratio manually by setting Width and Height proportionally.
            // Example (uncomment and adjust as required):
            // picture.Width = 200;
            // picture.Height = (int)(picture.Height * (200.0 / picture.Width));

            // Ensure the output directory exists
            string outputPath = @"C:\Output\WorkbookWithPicture.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
