// Title: C# – Insert an Image into Cell C9 and Retrieve Its Absolute X/Y Pixel Position with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a PNG picture to cell C9 (row 9, column C), set the picture to be placed inside the cell, read its absolute X and Y pixel coordinates, log the values, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# picture insertion | add image to Excel cell | cell C9 picture Aspose | absolute X coordinate Aspose.Cells | absolute Y coordinate Aspose.Cells | IsPlacedInCell property | .NET Excel image coordinates | worksheet.Pictures.Add example
// Common Searches: how to insert an image into a specific Excel cell with Aspose.Cells | retrieve pixel position of a picture placed in a cell using C# | Aspose.Cells get picture X Y coordinates | C# Aspose.Cells picture placement inside cell | absolute coordinates of Excel shape Aspose
// Developer Intent: Place a PNG image in cell C9 and obtain its exact pixel location for logging or further calculations.
// Use Cases: Validate image alignment when generating automated Excel reports. | Compute offsets for overlaying additional shapes relative to the inserted picture. | Export picture coordinates to external systems for UI mapping or analytics.
// AI Prompts: Generate C# code that adds a PNG to cell C9 with Aspose.Cells, sets IsPlacedInCell = true, and prints the picture's X and Y pixel values. | Show how to read and log the absolute X/Y coordinates of a shape after placing it inside a worksheet cell using Aspose.Cells for .NET. | Explain how to convert Aspose.Cells picture coordinates from points to pixels if needed.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add a PNG picture to cell C9 (row 9, column C), set the picture to be placed inside the cell, read its absolute X and Y pixel coordinates, log the values, and save the file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a picture to the worksheet at cell C9 (row index 8, column index 2)
        int pictureIndex = worksheet.Pictures.Add(8, 2, "sample.png");
        Picture picture = worksheet.Pictures[pictureIndex];

        // Place the picture inside the cell
        picture.IsPlacedInCell = true;

        // Retrieve the absolute X and Y coordinates (in pixels) of the picture
        double absoluteX = picture.X;
        double absoluteY = picture.Y;

        // Log the coordinates
        Console.WriteLine($"Picture absolute position - X: {absoluteX}, Y: {absoluteY}");

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
