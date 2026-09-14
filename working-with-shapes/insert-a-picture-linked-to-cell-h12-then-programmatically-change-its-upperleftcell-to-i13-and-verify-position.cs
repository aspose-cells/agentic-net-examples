// Title: Insert an image at cell H12 and reposition it to I13 with Aspose.Cells for .NET (C#)
// AI Prompts: Add a PNG picture anchored to cell H12 in a worksheet, then change its UpperLeftRow and UpperLeftColumn to anchor it at cell I13, and print the updated row and column. | Load an image file, insert it as a picture linked to H12 using Aspose.Cells, move the picture to I13 programmatically, and verify the new position by outputting the 1‑based row and column letter. | Create a new workbook, place an image at H12, adjust the picture's UpperLeftCell to I13, and save the workbook while confirming the picture's coordinates.
// Common Searches: Aspose.Cells C# insert picture at specific cell H12 and move to I13 | how to change picture UpperLeftRow UpperLeftColumn in Aspose.Cells | verify picture location after repositioning with Aspose.Cells .NET | programmatically reposition an image in Excel using Aspose.Cells C# | set picture anchor cell Aspose.Cells example
// Tags: insert picture Aspose.Cells C# | set picture UpperLeftRow UpperLeftColumn | move image to another cell Aspose.Cells | verify picture coordinates Aspose.Cells | Aspose.Cells picture positioning example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a PNG image linked to cell H12, updates the picture's UpperLeftRow and UpperLeftColumn to place it at cell I13, prints the new row (1‑based) and column letter, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "image.png";

            // Ensure the image file exists to avoid FileNotFoundException
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert picture linked to cell H12 (row index 11, column index 7)
            int startRow = 11;    // H12 -> row 12 (0‑based index)
            int startColumn = 7;  // H -> column 8 (0‑based index)
            int pictureIndex = sheet.Pictures.Add(startRow, startColumn, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Change the UpperLeftCell of the picture to I13 (row index 12, column index 8)
            picture.UpperLeftRow = 12;    // I13 -> row 13 (0‑based index)
            picture.UpperLeftColumn = 8; // I -> column 9 (0‑based index)

            // Verify the new position
            Console.WriteLine("Picture UpperLeftCell:");
            Console.WriteLine($"Row (1‑based): {picture.UpperLeftRow + 1}");
            Console.WriteLine($"Column (letter): {CellsHelper.ColumnIndexToName(picture.UpperLeftColumn)}");

            // Save the workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
