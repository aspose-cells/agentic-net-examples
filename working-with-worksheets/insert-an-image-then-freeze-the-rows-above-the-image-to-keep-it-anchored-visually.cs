// Title: Insert an Image and Freeze Rows Above It Using Aspose.Cells for .NET (C#)
// Description: Shows how to add a picture from a local file to a worksheet, apply FreezePanes to lock all rows above the picture, and save the result as an XLSX workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# image insertion | FreezePanes rows | add picture to Excel worksheet | anchor image in Excel | C# Aspose.Cells workbook save | Excel header logo freeze
// Common Searches: aspocells insert picture and freeze rows | c# freeze top rows after adding image | keep header logo visible in Excel using Aspose.Cells | freeze panes based on image position aspocells | add jpeg to worksheet and lock rows above
// Developer Intent: Place a picture in a worksheet and lock the rows above so the image remains visible while scrolling.
// Use Cases: Add a company logo at the top of a report that stays in view as users scroll through data. | Create a dashboard banner that is anchored to the first rows for consistent branding. | Generate a template with a watermark that should not move when the sheet is scrolled.
// AI Prompts: Generate C# code with Aspose.Cells that inserts a PNG at row 3 column 2 and freezes the first three rows. | Provide an example that adds a JPEG picture to a worksheet and uses FreezePanes to lock rows above the image in Aspose.Cells for .NET. | Show how to load an image from a file path, place it in a worksheet, and freeze rows up to the image's start row using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageFreezeDemo
{
    // Shows how to add a picture from a local file to a worksheet, apply FreezePanes to lock all rows above the picture, and save the result as an XLSX workbook with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the position where the image will be placed (top‑left cell)
                int imageTopRow = 5;      // Row index (0‑based) where the image starts
                int imageLeftColumn = 1; // Column index (0‑based) where the image starts

                // Resolve the image file path relative to the executable directory
                string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.jpg");

                // Add the picture if the file exists
                if (File.Exists(imagePath))
                {
                    worksheet.Pictures.Add(imageTopRow, imageLeftColumn, imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
                }

                // Freeze all rows above the image so it stays visually anchored
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Freeze up to the row where the image starts, no columns are frozen
                worksheet.FreezePanes(imageTopRow, 0, imageTopRow, 0);

                // Save the workbook
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ImageWithFrozenRows.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
