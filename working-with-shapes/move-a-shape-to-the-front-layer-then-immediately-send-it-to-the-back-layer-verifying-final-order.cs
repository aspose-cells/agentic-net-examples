// Title: Change a shape's Z-order: bring a shape to the front layer and then send it to the back using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to front layer | move Excel shape to back layer with Aspose.Cells .NET | verify shape order after changing Z-order in Aspose.Cells workbook | C# code to adjust shape layering in an Excel file using Aspose.Cells | how to bring a shape to front then back in Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition manipulation | C# adjust Excel shape layering | Aspose.Cells set shape front layer | Aspose.Cells set shape back layer | verify shape Z-order Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds two rectangle shapes, moves the first rectangle to the front by assigning its ZOrderPosition to the highest index, then immediately moves it to the back by setting ZOrderPosition to zero, checks that the shape is now the backmost item, prints the verification result, and saves the file as ShapeLayerOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (will be the first shape)
            Shape rectShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                1,   // upper left row
                1,   // upper left column
                0,   // top offset (in points)
                0,   // left offset (in points)
                100, // height (in points)
                50   // width (in points)
            );

            // Add a second rectangle shape (acts as a placeholder for the ellipse)
            Shape secondShape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // using rectangle since Ellipse is unavailable in this version
                2,   // upper left row
                1,   // upper left column
                0,   // top offset (in points)
                60,  // left offset (in points)
                100, // height (in points)
                50   // width (in points)
            );

            // Bring the first rectangle to the front layer
            rectShape.ZOrderPosition = sheet.Shapes.Count - 1;

            // Immediately send the first rectangle to the back layer
            rectShape.ZOrderPosition = 0;

            // Verify final order: the first rectangle should now be the backmost shape
            bool rectangleIsBack = sheet.Shapes[0] == rectShape;
            Console.WriteLine("Rectangle is at the back layer: " + rectangleIsBack);

            // Save the workbook to demonstrate persistence
            workbook.Save("ShapeLayerOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
