// Title: How to set a shape's shadow color using ARGB (255,128,0,0) and 30% transparency with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape shadow color using ARGB values | apply 30% transparency to shape shadow in Aspose.Cells for .NET | configure shadow effect angle blur distance for Excel shape with Aspose.Cells | how to add rectangle shape with custom shadow in a .NET workbook | set custom ARGB shadow color and transparency for Excel shape using Aspose.Cells
// Tags: shape shadow ARGB color Aspose.Cells | shadow transparency percentage Aspose.Cells | shadow effect angle blur distance Aspose.Cells | add rectangle shape with shadow Aspose.Cells | Excel shape shadow customization .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShadowEffectDemo
{
    // The example creates a workbook, adds a rectangle shape, and uses the Shape.ShadowEffect API to assign an ARGB color (255,128,0,0) and 30% transparency. It also demonstrates setting angle, blur, distance, and size before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to demonstrate the shadow effect
                // Parameters: upper left row, upper left column, upper left offset in pixels,
                //               upper left offset in pixels, width in points, height in points
                Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);

                // Access the shadow effect of the shape
                ShadowEffect shadow = shape.ShadowEffect;

                // Create a CellsColor and set its ARGB value to (255, 128, 0, 0)
                CellsColor shadowColor = workbook.CreateCellsColor();
                // ARGB packed into an int: (A << 24) | (R << 16) | (G << 8) | B
                shadowColor.Argb = (255 << 24) | (128 << 16) | (0 << 8) | 0;

                // Apply the color and transparency (30% = 0.3)
                shadow.Color = shadowColor;
                shadow.Transparency = 0.3; // 30% transparent

                // Optional: set additional shadow properties for better visibility
                shadow.Angle = 135;      // direction of the shadow
                shadow.Blur = 20;        // blur radius
                shadow.Distance = 10;    // distance from the shape
                shadow.Size = 1.0;       // size factor

                // Save the workbook
                string outputPath = "ShadowEffectOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Shadow color (ARGB): " + shadow.Color.Argb);
                Console.WriteLine("Shadow transparency: " + shadow.Transparency);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
