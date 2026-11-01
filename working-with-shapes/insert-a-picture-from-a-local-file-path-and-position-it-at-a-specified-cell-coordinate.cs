// Title: Insert a local PNG image into a specific worksheet cell and resize it using Aspose.Cells for .NET
// AI Prompts: Add a PNG from C:\Images\myPicture.png to cell B2 of a new workbook with Aspose.Cells and set its width to 200 px and height to 150 px. | Create an empty Excel file, place an image at row 1, column 1, adjust its size, and save the file to C:\Output\WorkbookWithPicture.xlsx. | Verify the image file exists, insert it into the first worksheet, modify the picture dimensions, and persist the workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to place an image at a specific cell address | resize inserted picture in Excel workbook using Aspose.Cells .NET | anchor PNG to B2 cell when generating Excel file with Aspose.Cells | programmatically add and scale a picture in a worksheet using Aspose.Cells for .NET
// Tags: add picture to worksheet cell Aspose.Cells | insert png image into Excel cell C# | set picture width and height Aspose.Cells | anchor image to specific cell coordinates Aspose.Cells | save workbook with embedded picture Aspose.Cells | verify image file existence before insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a local PNG file, inserts it into cell B2 of a newly created workbook, resizes the picture to 200 × 150 pixels, and saves the workbook to the designated output path.
class InsertPictureExample
{
    static void Main()
    {
        try
        {
            // Path to the image file to be inserted
            string imagePath = @"C:\Images\myPicture.png";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("Image file not found.", imagePath);

            // Path where the resulting Excel file will be saved
            string outputPath = @"C:\Output\WorkbookWithPicture.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Create a new workbook (empty Excel file)
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target cell coordinate (e.g., B2)
            // Row and column indexes are zero‑based, so B2 => row 1, column 1
            int targetRow = 1;    // B2 row
            int targetColumn = 1; // B2 column

            // Insert the picture; the top‑left corner of the picture will be anchored to the specified cell
            int pictureIndex = worksheet.Pictures.Add(targetRow, targetColumn, imagePath);

            // Optional: adjust picture size or offsets if needed
            Picture picture = worksheet.Pictures[pictureIndex];
            picture.Width = 200;   // width in pixels
            picture.Height = 150;  // height in pixels

            // Save the workbook to the specified file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
