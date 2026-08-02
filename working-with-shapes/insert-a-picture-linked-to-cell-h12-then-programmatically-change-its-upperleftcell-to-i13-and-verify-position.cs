// Title: C# – Insert a Picture at H12, Move to I13, and Verify Position with Aspose.Cells
// Description: Create a new workbook, add an image to cell H12, relocate its upper‑left corner to cell I13 using Picture.Move, output the new row/column indices, and save the file.
// Keywords: Aspose.Cells picture insertion | C# picture Move method | UpperLeftRow verification | Excel image placement programmatically | Aspose.Cells picture coordinates
// Common Searches: add image to specific cell Aspose.Cells C# | change picture UpperLeftCell Aspose.Cells | verify picture location after move Aspose.Cells | save workbook after moving picture Aspose.Cells
// Developer Intent: Programmatically place an image at H12, shift it to I13, and confirm the updated UpperLeftRow and UpperLeftColumn values.
// Use Cases: Insert a company logo into a template and reposition it when the layout changes. | Automate image alignment in dynamically generated financial reports. | Validate that pictures are correctly anchored after batch processing of Excel files.
// AI Prompts: Write C# code with Aspose.Cells to add a picture at H12, move it to I13, and print the new UpperLeftRow and UpperLeftColumn. | Explain how Picture.Move updates a picture's UpperLeftRow and UpperLeftColumn in Aspose.Cells. | Show error‑handling patterns for missing image files when inserting pictures with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Create a new workbook, add an image to cell H12, relocate its upper‑left corner to cell I13 using Picture.Move, output the new row/column indices, and save the file.
class InsertAndMovePicture
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

            if (File.Exists(imagePath))
            {
                // Add a picture at cell H12 (row index 11, column index 7)
                int pictureIndex = worksheet.Pictures.Add(11, 7, imagePath);
                Picture picture = worksheet.Pictures[pictureIndex];

                // Move the picture so its upper‑left corner is at cell I13 (row 12, column 8)
                picture.Move(12, 8);

                // Verify the new position
                Console.WriteLine($"Picture UpperLeftRow: {picture.UpperLeftRow}");
                Console.WriteLine($"Picture UpperLeftColumn: {picture.UpperLeftColumn}");
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Save the workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
