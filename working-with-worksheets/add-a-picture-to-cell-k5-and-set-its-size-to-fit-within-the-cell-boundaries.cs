// Title: Aspose.Cells for .NET: Insert and Fit an Image into Cell K5 (C#)
// Description: Creates a new workbook, checks for a JPEG file, adds the picture to cell K5, marks it as placed in the cell, and resizes it to match the column width and row height in pixels before saving the file.
// Keywords: Aspose.Cells add picture to cell | C# insert image Excel cell | fit picture to cell size | IsPlacedInCell property | resize picture Aspose.Cells
// Common Searches: how to embed an image in a specific Excel cell using Aspose.Cells | resize inserted picture to cell dimensions .NET | place picture inside a cell instead of floating | C# Aspose.Cells picture size matching cell
// Developer Intent: Add a picture to cell K5 and automatically size it to the cell's boundaries.
// Use Cases: Insert a company logo into a header cell that scales with the cell size. | Add product thumbnails to report cells without overlapping adjacent data. | Create a template where each cell contains a scaled icon or diagram.
// AI Prompts: Write C# code with Aspose.Cells to place an image in cell B2, set IsPlacedInCell to true, and adjust its Width and Height to the cell's pixel dimensions. | Show how to retrieve column width and row height in pixels and apply them to a picture object's size properties in Aspose.Cells. | Explain best practices for handling missing image files when inserting pictures into specific Excel cells with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsPictureInCell
{
    // Creates a new workbook, checks for a JPEG file, adds the picture to cell K5, marks it as placed in the cell, and resizes it to match the column width and row height in pixels before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the image file to be inserted
                string imagePath = "sample.jpg"; // replace with your image file

                // Ensure the image file exists
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Cell K5 corresponds to row index 4 (zero‑based) and column index 10 (zero‑based)
                int row = 4;      // K5 row
                int column = 10;  // K5 column

                // Add the picture so that its top‑left and bottom‑right corners are the same cell (K5)
                // This makes the picture occupy only that cell.
                int pictureIndex = sheet.Pictures.Add(row, column, row, column, imagePath);

                // Retrieve the inserted picture object
                Picture pic = sheet.Pictures[pictureIndex];

                // Ensure the picture is placed inside the cell (not floating over cells)
                pic.IsPlacedInCell = true;

                // Adjust the picture size to exactly match the cell dimensions
                pic.Width = sheet.Cells.GetColumnWidthPixel(column);
                pic.Height = sheet.Cells.GetRowHeightPixel(row);

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
}
