// Title: Aspose.Cells C# – Link Picture to a Cell, Set MoveAndSize Placement, Verify Resizing
// Description: Shows how to insert an image into a worksheet, bind it to cell B3, set Placement = MoveAndSize, resize row 3 and column B, and confirm that the picture automatically scales. Includes a fallback to an in‑memory PNG when the file is missing and saves the workbook as PictureMoveAndSize.xlsx.
// Keywords: Aspose.Cells picture linked cell | PlacementType.MoveAndSize | C# add image from file | C# add image from MemoryStream | image resize with cell Aspose.Cells | verify picture dimensions after cell resize | Aspose.Cells workbook save | row height column width affect picture | Aspose.Cells .NET image handling
// Common Searches: Aspose.Cells picture move and size with cell | link image to cell B3 Aspose.Cells | change row height column width affect picture size | add picture from MemoryStream Aspose.Cells C# | test picture resizing after cell resize Aspose.Cells
// Developer Intent: Insert an image, bind it to a specific cell, make it move and scale with that cell, and programmatically verify that adjusting the cell’s row height or column width updates the image size.
// Use Cases: Place a company logo in a header row that expands when the row height is increased. | Add placeholder graphics to a template where users can adjust column widths without breaking layout. | Generate data‑driven reports where pictures are attached to data cells and must follow cell size changes for consistent formatting.
// AI Prompts: Write C# code using Aspose.Cells to insert a picture from a file, link it to cell C5, set Placement = MoveAndSize, then change the row height and column width and output the new picture dimensions. | Show how to add a picture from a MemoryStream in Aspose.Cells, associate it with a cell, and confirm that the picture resizes when the cell’s dimensions are modified. | Explain the effect of PlacementType.MoveAndSize on picture behavior in Aspose.Cells and provide a sample test that validates automatic scaling after cell resizing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to insert an image into a worksheet, bind it to cell B3, set Placement = MoveAndSize, resize row 3 and column B, and confirm that the picture automatically scales. Includes a fallback to an in‑memory PNG when the file is missing and saves the workbook as PictureMoveAndSize.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "sample.jpg";

            int pictureIndex;

            // Add a picture to the worksheet.
            // If the file does not exist, use a minimal in‑memory PNG image.
            if (File.Exists(imagePath))
            {
                pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
            }
            else
            {
                // 1x1 pixel transparent PNG
                byte[] placeholderPng = new byte[]
                {
                    137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,
                    0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,
                    137,0,0,0,12,73,68,65,84,8,153,99,0,1,0,0,
                    5,0,1,13,10,2,0,0,0,0,73,69,78,68,174,66,
                    96,130
                };

                using (MemoryStream ms = new MemoryStream(placeholderPng))
                {
                    pictureIndex = worksheet.Pictures.Add(2, 2, ms);
                }
            }

            Picture picture = worksheet.Pictures[pictureIndex];

            // Link the picture to cell B3 and set placement to move and size with the cell
            picture.LinkedCell = "B3";
            picture.Placement = PlacementType.MoveAndSize;

            // Capture the initial size of the picture
            double initialHeight = picture.Height;
            double initialWidth = picture.Width;

            // Resize the linked cell (row 3 and column B)
            worksheet.Cells.SetRowHeight(2, 50);      // Row index 2 corresponds to row 3
            worksheet.Cells.SetColumnWidth(1, 30);   // Column index 1 corresponds to column B

            // After resizing the cell, capture the new size of the picture
            double newHeight = picture.Height;
            double newWidth = picture.Width;

            // Output size information
            Console.WriteLine($"Initial size - Height: {initialHeight}, Width: {initialWidth}");
            Console.WriteLine($"After cell resize - Height: {newHeight}, Width: {newWidth}");

            // Save the workbook
            workbook.Save("PictureMoveAndSize.xlsx", SaveFormat.Xlsx);
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
