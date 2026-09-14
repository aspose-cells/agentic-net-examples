// Title: Measure XLSX workbook size before and after resizing a rectangle shape with Aspose.Cells for .NET
// AI Prompts: Create a C# console program that adds a rectangle shape to a worksheet, saves the workbook to a MemoryStream, changes the shape's width, height, and text, saves again, and prints the original size, new size, and byte difference. | Implement a reusable method `GetWorkbookSizeDelta(Workbook wb, Action<Shape> modify)` that accepts a workbook and a shape‑modification delegate, saves the workbook twice to MemoryStream, and returns the size delta. | Add comprehensive error handling that catches Aspose.Cells exceptions during shape operations and still reports the size difference.
// Common Searches: aspnet measure Excel file size change after editing drawing objects with Aspose.Cells | c# get byte length of saved XLSX workbook using MemoryStream Aspose | how much does resizing a shape affect XLSX file size in Aspose.Cells | compare original and modified workbook sizes after shape manipulation in .NET | calculate size delta of Excel file after changing rectangle dimensions with Aspose.Cells
// Tags: calculate XLSX size delta after shape resize | Aspose.Cells save workbook to MemoryStream for size measurement | measure workbook file size change .NET | rectangle shape dimension modification impact on Excel size | compare original and modified Excel file size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a workbook, inserts a rectangle shape, saves it to a MemoryStream, modifies the shape's dimensions and text, saves again, and outputs the original size, modified size, and the byte difference.
class ShapeSizeComparison
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wbOriginal = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wbOriginal.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);
            rect.Text = "Original Shape";

            // Save the original workbook to a memory stream to measure its size
            using (MemoryStream msOriginal = new MemoryStream())
            {
                wbOriginal.Save(msOriginal, SaveFormat.Xlsx);
                long originalSize = msOriginal.Length;

                // Modify the shape: change its size and text
                rect.Width = 300;   // increase width
                rect.Height = 150;  // increase height
                rect.Text = "Modified Shape";

                // Save the modified workbook to another memory stream
                using (MemoryStream msModified = new MemoryStream())
                {
                    wbOriginal.Save(msModified, SaveFormat.Xlsx);
                    long modifiedSize = msModified.Length;

                    // Calculate size difference
                    long sizeDifference = modifiedSize - originalSize;

                    // Output the results
                    Console.WriteLine($"Original workbook size: {originalSize} bytes");
                    Console.WriteLine($"Modified workbook size: {modifiedSize} bytes");
                    Console.WriteLine($"Size difference: {sizeDifference} bytes");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
