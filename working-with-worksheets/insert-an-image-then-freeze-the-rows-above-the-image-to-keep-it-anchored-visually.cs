// Title: Insert a PNG image into an Excel worksheet and freeze rows above it using Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a PNG picture to cell A1 of a worksheet and then freezes every row above the picture using Aspose.Cells. | Create a method that reads a picture’s UpperLeftRow and applies FreezePanes so the image stays visible while scrolling.
// Common Searches: Aspose.Cells C# insert image at A1 and freeze rows above picture | how to keep an inserted picture anchored while scrolling in Excel with Aspose.Cells | programmatically apply FreezePanes based on picture position in a .NET workbook | freeze rows above a picture using Aspose.Cells FreezePanes method
// Tags: add PNG picture to worksheet Aspose.Cells | apply FreezePanes based on picture UpperLeftRow C# | anchor image with frozen rows Aspose.Cells | picture positioning and row freezing .NET | insert image and lock rows in Excel using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a new workbook, inserts a PNG image at cell A1, determines the picture's UpperLeftRow, freezes all rows above that row with FreezePanes, and saves the file as output.xlsx.
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
            string imagePath = "sample.png";

            // Ensure the image file exists before attempting to insert
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the image at cell A1 (row 0, column 0)
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Determine the row index where the picture starts (UpperLeftRow)
            // Freeze all rows above this row so the image stays anchored visually
            int freezeRow = picture.UpperLeftRow; // rows with index < freezeRow will be frozen

            // Apply freeze panes: freeze rows above freezeRow, no columns are frozen
            // Parameters: (freezeRows, freezeColumns, topRow, leftColumn)
            sheet.FreezePanes(freezeRow, 0, freezeRow, 0);

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
