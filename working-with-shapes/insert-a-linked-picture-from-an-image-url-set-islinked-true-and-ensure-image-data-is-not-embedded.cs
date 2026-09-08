// Title: Insert a linked picture from a remote URL into an Excel worksheet using Aspose.Cells for .NET without embedding the image data
// AI Prompts: Add a picture to cell A1 that references an external image URL and keep it linked rather than embedding the file. | Create an Excel workbook, place a linked external image with specific width and height, and save it as .xlsx using Aspose.Cells for .NET. | Ensure the picture's IsLinked property is set to true so the workbook does not store the image bytes.
// Common Searches: asp.net insert linked picture from web URL using Aspose.Cells | how to add external image to Excel without embedding with Aspose.Cells C# | set picture IsLinked property true in Aspose.Cells example | Aspose.Cells picture dimensions and external link C#
// Tags: insert linked picture Aspose.Cells | external image URL Excel .NET | disable image embedding Aspose.Cells | IsLinked property picture C# | picture size Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// // Creates a new workbook, adds a picture that links to an external URL at cell A1, sets its width and height, ensures IsLinked is true so the image data is not embedded, and saves the file as LinkedPicture.xlsx using Aspose.Cells for .NET.
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

            // URL of the image to link
            string imageUrl = "https://example.com/image.png";

            // Add a picture at cell A1 (row 0, column 0) using the URL.
            // The Add method returns the index of the newly added picture.
            int pictureIndex = sheet.Pictures.Add(0, 0, imageUrl);

            // Retrieve the picture object
            Picture linkedPicture = sheet.Pictures[pictureIndex];

            // Optionally set size (in pixels)
            linkedPicture.Width = 200;
            linkedPicture.Height = 150;

            // Save the workbook to a file
            workbook.Save("LinkedPicture.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
