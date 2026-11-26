// Title: Calculate XLSX file size difference before and after modifying a rectangle shape with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get workbook size before and after changing shape dimensions | how to measure Excel file size delta after updating shape fill color using Aspose.Cells | C# compare memory stream lengths of XLSX files with modified rectangle shape | determine size impact of shape property changes in an Aspose.Cells workbook | retrieve original and modified workbook byte size in .NET Aspose.Cells example
// Tags: shape property change file size Aspose.Cells | workbook memory stream size measurement C# | rectangle shape dimension impact on XLSX size | calculate workbook size delta Aspose.Cells | modify shape fill color size effect

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeSizeComparisonApp
{
    // The sample creates a workbook, adds a rectangle shape, saves it to a memory stream to capture the original XLSX size, then alters the shape's text, dimensions, and fill color, saves the workbook again, and computes the byte‑size difference between the original and modified files.
    class ShapeSizeComparison
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 200);
                rect.Text = "Original Shape";

                // Save the workbook to a memory stream to get the original size
                using (MemoryStream originalStream = new MemoryStream())
                {
                    workbook.Save(originalStream, SaveFormat.Xlsx);
                    long originalSize = originalStream.Length;

                    // Modify the shape: change text, size, and fill color
                    rect.Text = "Modified Shape";
                    rect.Width = 300;   // increase width
                    rect.Height = 150;  // change height
                    // Set fill color using FillFormat
                    rect.FillFormat.ForeColor = Color.LightBlue;

                    // Save the modified workbook to another memory stream
                    using (MemoryStream modifiedStream = new MemoryStream())
                    {
                        workbook.Save(modifiedStream, SaveFormat.Xlsx);
                        long modifiedSize = modifiedStream.Length;

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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
