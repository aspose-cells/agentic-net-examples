// Title: Add a picture to cell K5 and fit it within the cell using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, verifies a JPEG file, inserts the image into cell K5, sets IsPlacedInCell to keep the picture inside the cell boundaries, and saves the file as PictureInK5.xlsx.
// Keywords: Aspose.Cells add picture C# | insert image into worksheet cell | IsPlacedInCell property | fit picture to cell boundaries | Aspose.Cells picture placement
// Common Searches: Aspose.Cells add image to specific cell | C# place picture inside Excel cell | fit picture to cell size Aspose.Cells | set IsPlacedInCell true Aspose.Cells | insert JPEG into Excel cell using Aspose
// Developer Intent: Insert a JPEG into cell K5 and ensure the image remains confined to the cell's dimensions.
// Use Cases: Embedding a company logo in a designated cell of a financial report. | Displaying product thumbnails in inventory worksheets. | Adding a signature image to a form cell for electronic approval.
// AI Prompts: Write C# code with Aspose.Cells that inserts an image into cell K5 and automatically scales it to the cell size. | Show how to adjust picture height and width after enabling IsPlacedInCell in Aspose.Cells. | Provide robust error handling for missing image files when adding a picture to a worksheet with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, verifies a JPEG file, inserts the image into cell K5, sets IsPlacedInCell to keep the picture inside the cell boundaries, and saves the file as PictureInK5.xlsx.
class AddPictureToCell
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "sample.jpg";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to cell K5 (row 5, column K)
            // Row index for row 5 is 4 (zero‑based), column index for column K is 10.
            // Using the overload that specifies topRow, leftColumn, bottomRow, rightColumn, and file name.
            int pictureIndex = worksheet.Pictures.Add(4, 10, 4, 10, imagePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Ensure the picture is placed inside the cell boundaries
            picture.IsPlacedInCell = true;

            // Save the workbook
            string outputPath = "PictureInK5.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
