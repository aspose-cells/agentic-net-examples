// Title: Lock Excel shapes and verify lock enforcement when moving a shape with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock a shape to prevent movement | how to test shape lock behavior in an Excel workbook using Aspose.Cells | programmatically move a locked shape and catch exception with Aspose.Cells | set IsLocked property for shapes in Aspose.Cells .NET example | verify that a locked shape cannot be repositioned in Aspose.Cells
// Tags: Aspose.Cells shape locking C# | IsLocked property for Excel shapes .NET | prevent shape movement with Aspose.Cells | shape lock enforcement example | testing shape lock programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds two rectangle shapes, sets their IsLocked property to true, then attempts to reposition one shape inside a try‑catch block to demonstrate that the lock prevents movement, and finally saves the file as ShapeLockTest.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape
            Shape rect = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // row offset (in pixels)
                0,   // column offset (in pixels)
                100, // height (in points)
                50   // width (in points)
            );
            rect.Text = "Rectangle";

            // Add a second shape (using rectangle as fallback for ellipse)
            Shape secondShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // fallback shape type
                5,
                5,
                0,
                0,
                80,
                80
            );
            secondShape.Text = "Second Shape";

            // Lock both shapes to prevent editing in the UI
            rect.IsLocked = true;
            secondShape.IsLocked = true;

            // Attempt to move the first shape programmatically
            try
            {
                rect.Left = 200; // move horizontally
                rect.Top = 200;  // move vertically
                Console.WriteLine("Rectangle moved despite lock.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Cannot move rectangle because it is locked: " + ex.Message);
            }

            // Save the workbook to a file
            string outputPath = "ShapeLockTest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
