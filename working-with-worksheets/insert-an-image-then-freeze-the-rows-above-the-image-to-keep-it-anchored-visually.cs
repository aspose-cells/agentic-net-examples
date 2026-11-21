// Title: Insert a picture into an Excel worksheet and freeze the rows above it with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells insert image at specific cell and freeze top rows | How to use FreezePanes after adding a picture in Aspose.Cells .NET | Aspose.Cells picture insertion with row freeze for Excel workbook | Freeze rows above an inserted image using Aspose.Cells in C# | Add PNG to worksheet and keep it anchored by freezing rows with Aspose.Cells
// Tags: insert picture worksheet Aspose.Cells C# | freeze panes rows above image Aspose.Cells | anchor image with FreezePanes .NET | add PNG to Excel using Aspose.Cells | programmatic row freeze after picture insertion

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, checks for an image file, inserts the PNG at row 5 column A, freezes the rows above the picture using the FreezePanes method, and saves the workbook as an XLSX file, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image to insert
            string imagePath = "image.png";

            // Verify that the image file exists before attempting to add it
            if (File.Exists(imagePath))
            {
                // Insert the image. Place its top‑left corner at row index 4 (5th row) and column index 0 (first column)
                int pictureRow = 4;      // zero‑based row index
                int pictureColumn = 0;   // zero‑based column index
                sheet.Pictures.Add(pictureRow, pictureColumn, imagePath);

                // Freeze rows above the inserted image (rows 0‑4) and no columns
                // FreezePanes(row, column, totalRows, totalColumns)
                sheet.FreezePanes(pictureRow + 1, 0, pictureRow + 1, 0);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping image insertion.");
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
