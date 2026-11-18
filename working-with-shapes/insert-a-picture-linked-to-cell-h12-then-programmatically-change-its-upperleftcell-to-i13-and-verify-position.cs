// Title: Insert a picture anchored to cell H12 and move it to cell I13 using Aspose.Cells for .NET
// Common Searches: how to anchor an image to a specific cell using Aspose.Cells C# | change picture UpperLeftRow and UpperLeftColumn to reposition image in Aspose.Cells workbook | verify picture anchor cell after moving image with Aspose.Cells .NET API | programmatically move a linked picture from H12 to I13 in Excel using Aspose.Cells | Aspose.Cells picture positioning example C#
// Tags: insert picture anchored to cell Aspose.Cells | set picture UpperLeftRow UpperLeftColumn Aspose.Cells | reposition picture to another cell Aspose.Cells | retrieve picture anchor coordinates Aspose.Cells | Aspose.Cells picture positioning C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, inserts a picture linked to cell H12, updates its UpperLeftRow and UpperLeftColumn to point to cell I13, prints the new cell reference to confirm the move, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image to be inserted
            string imagePath = "sample.png";

            // Ensure the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert a picture linked to cell H12 (row 11, column 7 - zero based indices)
            int rowH12 = 11; // H12 row index
            int colH12 = 7;  // H column index
            int pictureIdx = sheet.Pictures.Add(rowH12, colH12, imagePath);
            Picture picture = sheet.Pictures[pictureIdx];

            // Change the UpperLeftCell of the picture to I13 (row 12, column 8)
            picture.UpperLeftRow = 12;    // I13 row index
            picture.UpperLeftColumn = 8; // I column index

            // Verify the new position
            string upperLeftCellName = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);
            Console.WriteLine($"Picture UpperLeftCell is now: {upperLeftCellName}"); // Expected output: I13

            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
