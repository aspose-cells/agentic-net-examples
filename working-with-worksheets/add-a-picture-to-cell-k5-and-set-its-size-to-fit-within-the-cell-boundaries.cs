// Title: Insert an image into cell K5 and automatically resize it to fit the cell using Aspose.Cells for .NET
// AI Prompts: Add a PNG picture to worksheet cell K5 and set its Width and Height to the cell's pixel dimensions with Aspose.Cells in C#. | Retrieve the pixel width of column K and the pixel height of row 5, then apply those values to a Picture object so the image fills the cell. | Programmatically place an image in a specific Excel cell and scale it to the cell boundaries using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# insert image into specific cell and fit to cell size | Resize picture to match Excel cell dimensions using Aspose.Cells .NET | Get column width in pixels Aspose.Cells and set picture size | How to add a PNG to cell K5 and auto‑scale it with Aspose.Cells | Fit an image inside an Excel cell programmatically with Aspose.Cells
// Tags: Aspose.Cells add image to specific cell | Aspose.Cells resize picture to cell size | Aspose.Cells get column width in pixels | Aspose.Cells get row height in pixels | C# scale image to fit Excel cell | Aspose.Cells picture width height properties

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, verifies the presence of the input Excel file and PNG image, inserts the image into cell K5, obtains the cell's pixel width and height, resizes the picture to those dimensions, and saves the result, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string imagePath = "image.png";
            const string outputPath = "output.xlsx";

            // Verify required files exist to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Target cell K5 (zero‑based indices: row 4, column 10)
            int targetRow = 4;
            int targetColumn = 10;

            // Add the picture to the worksheet at the target cell
            int pictureIndex = sheet.Pictures.Add(targetRow, targetColumn, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Retrieve the cell's width and height in pixels
            double cellWidth = sheet.Cells.GetColumnWidthPixel(targetColumn);
            double cellHeight = sheet.Cells.GetRowHeightPixel(targetRow);

            // Resize the picture to fit within the cell boundaries
            picture.Width = (int)cellWidth;
            picture.Height = (int)cellHeight;

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
