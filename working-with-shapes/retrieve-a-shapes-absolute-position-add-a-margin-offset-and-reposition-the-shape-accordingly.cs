// Title: Read a shape’s absolute Top/Left coordinates and shift it by a margin offset with Aspose.Cells for .NET
// Common Searches: aspnet cells get shape top left coordinates c# | aspnet cells move an Excel shape by a specific offset c# | aspnet cells add margin to a worksheet shape programmatically | aspnet cells reposition shape using Top and Left properties c# | aspnet cells adjust shape location after loading workbook c#
// Tags: Aspose.Cells shape repositioning | Aspose.Cells shape coordinate offset | Aspose.Cells get shape absolute location | Aspose.Cells worksheet shape margin adjustment | Aspose.Cells C# shape positioning

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads an Excel workbook, reads the Top and Left properties of the first shape on the first worksheet, adds a 10‑point horizontal and 5‑point vertical margin, updates the shape’s position, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape (or use sheet.Shapes["ShapeName"] for a named shape)
            Shape shape = sheet.Shapes[0];

            // Get current absolute position (in points)
            int currentTop = shape.Top;
            int currentLeft = shape.Left;

            // Define margin offsets (in points)
            int offsetX = 10; // shift right by 10 points
            int offsetY = 5;  // shift down by 5 points

            // Compute new position
            int newTop = currentTop + offsetY;
            int newLeft = currentLeft + offsetX;

            // Apply new position to the shape
            shape.Top = newTop;
            shape.Left = newLeft;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
