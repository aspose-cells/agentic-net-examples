// Title: Insert a PNG picture into cell A6 and freeze the rows above it with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to place a PNG image at cell A6 and then freezes the first five rows so the picture stays visible while scrolling. | Create a .NET example that adds a picture to a worksheet, anchors it by freezing the rows above the image, and saves the workbook. | Show how to verify the image file exists, insert it, and apply FreezePanes in Aspose.Cells C#.
// Common Searches: asp.net insert PNG into specific cell and freeze rows above with Aspose.Cells | C# Aspose.Cells freeze panes after adding a picture to Excel worksheet | keep inserted image fixed while scrolling in Excel using Aspose.Cells .NET | example code to anchor a picture by freezing top rows in Aspose.Cells | how to use FreezePanes to lock rows above an inserted image in C#
// Tags: insert picture into worksheet Aspose.Cells | freeze top rows using FreezePanes Aspose.Cells | anchor PNG image in Excel .NET | Aspose.Cells picture placement at cell A6 | freeze rows above inserted image C# | save workbook with image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, inserts a PNG picture at cell A6, freezes the first five rows so the image remains visible while scrolling, and saves the file as ImageWithFrozenRows.xlsx.
class InsertImageAndFreezeRows
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = @"C:\Images\SampleImage.png";

            // Verify that the image file exists
            if (!File.Exists(imagePath))
                throw new FileNotFoundException("Image file not found.", imagePath);

            // Insert the image at cell A6 (row index 5, column index 0)
            int pictureRow = 5;      // zero‑based row index (A6)
            int pictureColumn = 0;   // zero‑based column index (A)
            sheet.Pictures.Add(pictureRow, pictureColumn, imagePath);

            // Freeze the first 5 rows (rows 0‑4) so the image stays anchored visually.
            // Use the overload that specifies the number of rows and columns to freeze.
            sheet.FreezePanes(pictureRow, 0, pictureRow, 0);

            // Ensure the output directory exists
            string outputPath = @"C:\Output\ImageWithFrozenRows.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook to a file
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
