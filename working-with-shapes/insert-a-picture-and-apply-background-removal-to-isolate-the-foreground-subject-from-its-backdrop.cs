// Title: Insert a picture into an Excel worksheet and remove its background with Aspose.Cells for .NET
// Description: Demonstrates how to add an image to a worksheet, set a transparent color to hide the white backdrop, optionally resize the picture, and save the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells insert image | Excel picture transparent background .NET | remove white background Aspose.Cells | picture formatting Aspose.Cells | C# Aspose.Cells image resizing | background removal Excel .NET | transparent color picture Aspose
// Common Searches: Aspose.Cells make picture background transparent | C# insert image into Excel with transparent background | remove white background from picture in Excel using Aspose | how to set TransparentColor for a picture in Aspose.Cells | resize inserted image Aspose.Cells C#
// Developer Intent: Add an image to a worksheet and make its background transparent programmatically.
// Use Cases: Embedding a logo in a financial report without the surrounding white margin. | Displaying product photos in a catalog sheet that blend with cell colors. | Creating dashboards where overlay images do not obscure underlying data.
// AI Prompts: Write C# code that inserts a PNG into an Aspose.Cells worksheet and sets TransparentColor to a specific RGB value. | Show how to batch‑process a folder of JPEGs, add each to a separate worksheet, apply white‑background removal, and adjust dimensions with Aspose.Cells. | Explain how to detect the dominant background hue of an image and automatically apply it as TransparentColor before insertion.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an image to a worksheet, set a transparent color to hide the white backdrop, optionally resize the picture, and save the workbook as an XLSX file using Aspose.Cells for .NET.
class InsertPictureWithBackgroundRemoval
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the picture file to be inserted
            string picturePath = "input.jpg";

            // Ensure the picture file exists before adding
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Picture file not found: {picturePath}");

            // Add the picture to the worksheet at cell (0,0)
            int pictureIndex = worksheet.Pictures.Add(0, 0, picturePath);

            // Retrieve the Picture object
            Picture picture = worksheet.Pictures[pictureIndex];

            // Make white background transparent
            CellsColor transparentColor = workbook.CreateCellsColor();
            transparentColor.Color = Color.White;
            picture.FormatPicture.TransparentColor = transparentColor;

            // Adjust picture size if needed
            picture.Width = 300;   // width in pixels
            picture.Height = 200;  // height in pixels

            // Save the workbook
            workbook.Save("OutputWithTransparentBackground.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
