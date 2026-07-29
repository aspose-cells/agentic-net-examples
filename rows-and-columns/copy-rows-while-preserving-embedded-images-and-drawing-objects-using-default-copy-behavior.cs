// Title: Copy Rows with Images & Drawing Objects – Aspose.Cells for .NET (C#)
// Description: Use the Cells.CopyRows method to duplicate a range of rows while automatically preserving embedded pictures, charts, and other drawing objects. This C# example creates a workbook, adds sample data, inserts an image, copies the rows to a new location, and saves the result, demonstrating the default copy behavior in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | copy rows | preserve images | drawing objects | Cells.CopyRows | Excel automation | .NET | embedded picture | worksheet duplication
// Common Searches: Aspose.Cells copy rows with picture | C# copy Excel rows keep images | How to duplicate rows with charts Aspose.Cells | Copy rows preserving drawing objects .NET | Cells.CopyRows example with embedded images
// Developer Intent: Duplicate a set of rows in an Excel worksheet while keeping any embedded pictures, charts, or shapes intact using Aspose.Cells for .NET.
// Use Cases: Copy a header row that contains a company logo to another section of the report without losing the logo. | Replicate data rows that include icons or sparklines for each record in a generated invoice. | Create a template where rows with embedded charts are copied multiple times to build a multi‑page dashboard.
// AI Prompts: Show C# code that copies rows with embedded pictures using Aspose.Cells and adjusts picture positions after the copy. | Provide an example of copying rows that contain drawing objects and then renaming the destination worksheet. | Explain how Cells.CopyRows handles embedded images and what, if any, additional settings are required to ensure they are preserved.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Use the Cells.CopyRows method to duplicate a range of rows while automatically preserving embedded pictures, charts, and other drawing objects. This C# example creates a workbook, adds sample data, inserts an image, copies the rows to a new location, and saves the result, demonstrating the default copy behavior in Aspose.Cells for .NET.
class CopyRowsWithImagesDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook (source)
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Populate some sample data in the first three rows
            sourceSheet.Cells["A1"].PutValue("Row 1");
            sourceSheet.Cells["A2"].PutValue("Row 2");
            sourceSheet.Cells["A3"].PutValue("Row 3");

            // Insert an image (embedded picture) that resides in the second row
            string imagePath = "sample_image.png";
            if (File.Exists(imagePath))
            {
                int pictureIndex = sourceSheet.Pictures.Add(1, 1, imagePath);
                Picture picture = sourceSheet.Pictures[pictureIndex];
                // Optionally set picture properties (size, position) here
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Copy all existing rows (including the image) to a new location starting at row 5
            sourceSheet.Cells.CopyRows(
                sourceSheet.Cells,          // source cells
                0,                         // source start row index
                5,                         // destination start row index
                sourceSheet.Cells.MaxDisplayRange.RowCount // number of rows to copy
            );

            // Save the workbook to verify that rows and the embedded image were copied
            string outputPath = "CopyRowsWithImages.xlsx";
            sourceWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
