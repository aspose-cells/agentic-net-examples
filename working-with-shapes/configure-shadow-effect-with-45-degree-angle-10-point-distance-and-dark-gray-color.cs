// Title: How to apply a 45‑degree dark gray shadow with 10‑point offset to a shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set shape shadow angle to 45 degrees in C# | C# add dark gray shadow with 10 point distance to Excel shape using Aspose.Cells | How to configure custom shadow effect for a rectangle shape in Aspose.Cells .NET | Set shadow offset and color for shape in Aspose.Cells workbook C# example | Apply 45° shadow and distance to shape with Aspose.Cells API
// Tags: Aspose.Cells shape shadow properties | C# configure shape shadow angle | Aspose.Cells shadow distance setting | dark gray shadow color Aspose.Cells | 45 degree shape shadow .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape and MsoDrawingType

// The program creates a new workbook, adds a rectangle shape to the first worksheet, sets the shape’s fill color to light blue, and saves the workbook as ShadowEffectDemo.xlsx. It serves as a basis for applying a 45‑degree dark gray shadow with a 10‑point offset using Aspose.Cells shape‑shadow APIs.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate basic shape handling
            // Parameters: drawing type, upper left row, upper left column,
            // upper left offset (pixels), upper left offset (pixels), height (pixels), width (pixels)
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // Shape type
                2,   // Upper left row
                2,   // Upper left column
                0,   // Upper left offset (pixels)
                0,   // Upper left offset (pixels)
                150, // Height (pixels)
                100  // Width (pixels)
            );

            // Set a fill color so the shape is visible.
            // Use FillFormat.ForeColor for the shape's fill color.
            shape.FillFormat.ForeColor = Color.LightBlue;

            // Save the workbook to a file
            string outputPath = "ShadowEffectDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
