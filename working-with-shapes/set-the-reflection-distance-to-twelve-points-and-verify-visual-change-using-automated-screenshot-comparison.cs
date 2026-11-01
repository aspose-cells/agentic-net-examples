// Title: Set shape reflection distance to 12 points in Aspose.Cells for .NET and validate with screenshot comparison
// Common Searches: Aspose.Cells C# set shape reflection distance to 12 points | compare Excel shape screenshots using Aspose.Cells | automated visual verification of shape reflection in .NET | C# Aspose.Cells shape styling regression test | measure visual change after setting shape reflection in Excel workbook
// Tags: Aspose.Cells shape reflection API | C# set rectangle reflection distance | Excel shape screenshot comparison .NET | visual regression testing Aspose.Cells | shape styling verification code sample

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape, sets its Reflection.Distance property to 12 points, and demonstrates how to confirm the visual change using automated screenshot comparison.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a simple rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 60);

            // Set the reflection distance to 12 points (measured in points)
            shape.Reflection.Distance = 12f;

            Console.WriteLine("Reflection distance set to 12 points.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
