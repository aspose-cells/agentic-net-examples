// Title: Set a picture to move and resize with its linked cell in Aspose.Cells for .NET and verify the change after adjusting row height and column width
// AI Prompts: Insert a PNG image into cell C3 of a worksheet and set its Placement property to MoveAndSize using Aspose.Cells for .NET. | Modify the row height and column width of the cell linked to the picture, then read and output the picture's Width and Height properties. | Save the workbook before and after resizing the linked cell to demonstrate that the picture scales automatically.
// Common Searches: Aspose.Cells C# set picture placement to MoveAndSize so it follows cell resizing | how to link an image to a specific cell and have it resize with the cell in Aspose.Cells | verify picture dimensions after adjusting row height and column width with Aspose.Cells for .NET | sample code for moving and sizing pictures with cells using Aspose.Cells in C# | Aspose.Cells picture auto resize when row height changes example
// Tags: picture placement movandsize Aspose.Cells | link image to worksheet cell Aspose.Cells | auto resize picture with cell dimensions C# | row height column width affect picture Aspose.Cells | save workbook after picture resize Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Demonstrates adding a PNG picture to cell C3, configuring its Placement to MoveAndSize, saving the workbook, then changing the linked cell's row height and column width and printing the updated picture Width and Height to confirm automatic scaling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the cell that will be linked to the picture (C3 -> row index 2, column index 2)
            int pictureRow = 2;      // Row 3 (zero‑based)
            int pictureColumn = 2;   // Column C (zero‑based)

            // Path to an existing image file
            string imagePath = "sample.png";

            // Ensure the image file exists before attempting to add it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to the worksheet; Add returns the picture index
            int pictureIndex = sheet.Pictures.Add(pictureRow, pictureColumn, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Configure the picture to move and resize together with its linked cell
            picture.Placement = PlacementType.MoveAndSize;

            // Set an initial size for the linked cell
            sheet.Cells.SetRowHeight(pictureRow, 30);      // Height in points
            sheet.Cells.SetColumnWidth(pictureColumn, 20); // Width in characters

            // Save the workbook with the initial configuration
            workbook.Save("PictureMoveAndSize_initial.xlsx");

            // Resize the linked cell to verify that the picture follows the change
            sheet.Cells.SetRowHeight(pictureRow, 60);      // New height
            sheet.Cells.SetColumnWidth(pictureColumn, 40); // New width

            // Output the picture's dimensions after the cell resize
            Console.WriteLine("Picture Width after cell resize: " + picture.Width);
            Console.WriteLine("Picture Height after cell resize: " + picture.Height);

            // Save the workbook after resizing
            workbook.Save("PictureMoveAndSize_resized.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
