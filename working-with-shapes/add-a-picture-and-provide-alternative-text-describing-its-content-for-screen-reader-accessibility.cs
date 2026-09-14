// Title: Add a PNG picture to cell B2 in an Excel workbook and define AlternativeText for accessibility using Aspose.Cells for .NET
// AI Prompts: Insert a PNG image from a file stream into cell B2 of a new workbook and set its AlternativeText property with Aspose.Cells in C#. | Create a workbook, add a picture shape at a specific cell, assign descriptive alt text for screen readers, and save the file using the Aspose.Cells .NET API. | Load an image file, place it as a picture on the first worksheet, configure the picture's AlternativeText for accessibility, and export the workbook.
// Common Searches: how to set alt text for an inserted picture in Aspose.Cells C# | Aspose.Cells add image to specific cell with alternative description | C# example for adding PNG to Excel and providing accessibility text | using Aspose.Cells to insert picture and define AlternativeText property | save workbook with picture and alt text using Aspose.Cells for .NET
// Tags: add picture shape to worksheet Aspose.Cells C# | configure picture AlternativeText Aspose.Cells | place image at specific cell Aspose.Cells | accessibility description for Excel picture Aspose.Cells | picture properties manipulation Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, loads a PNG file, inserts it as a picture at cell B2, assigns an AlternativeText description for screen‑reader accessibility, and saves the workbook as Workbook_With_Picture.xlsx.
class AddPictureWithAltText
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Path to the image file to be inserted
        string imagePath = "sample_image.png";

        // Ensure the image file exists
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }

        // Load the image into a stream
        using (FileStream imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        {
            // Add the picture to the worksheet at cell B2 (row index 1, column index 1)
            // Lifecycle rule: create picture via Pictures.Add
            int pictureIndex = sheet.Pictures.Add(1, 1, imageStream);

            // Retrieve the inserted picture object
            Picture picture = sheet.Pictures[pictureIndex];

            // Set alternative text for screen reader accessibility
            picture.AlternativeText = "A scenic view of mountains during sunrise";

            // Optionally, adjust picture size or position here if needed
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save("Workbook_With_Picture.xlsx");
    }
}
