using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a picture to cell C9 (row index 8, column index 2)
        int pictureIndex = worksheet.Pictures.Add(8, 2, "sample.png");
        Picture picture = worksheet.Pictures[pictureIndex];

        // Place the picture inside the cell
        picture.IsPlacedInCell = true;

        // Retrieve absolute X and Y coordinates (in pixels)
        double absoluteX = picture.X;
        double absoluteY = picture.Y;

        // Log the coordinates
        Console.WriteLine($"Picture absolute position - X: {absoluteX}, Y: {absoluteY}");

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}