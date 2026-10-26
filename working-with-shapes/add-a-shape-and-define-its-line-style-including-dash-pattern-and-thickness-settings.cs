// Title: Create a rectangle shape with a 2.5‑point dashed border in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to add a rectangle shape with a custom dashed border in Aspose.Cells C# | Aspose.Cells set shape line thickness and dash style programmatically | C# code to draw a shape with specific line weight and dash pattern in Excel | Aspose.Cells shape line formatting example for .NET developers | set shape border dash style using Aspose.Cells workbook
// Tags: Aspose.Cells rectangle shape line formatting | C# set shape dash style Excel | Aspose.Cells line weight property | Excel shape border thickness .NET | Aspose.Cells shape line dash pattern

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, sets the shape's line weight to 2.5 points and its dash style to a dashed pattern, optionally assigns a line color, and saves the file as ShapeWithLineStyle.xlsx.
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

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 120);

            // Set the line (border) thickness (points)
            shape.Line.Weight = 2.5;

            // Set the dash pattern for the line
            shape.Line.DashStyle = MsoLineDashStyle.Dash;

            // Optional: set line color (commented out if not supported by the current library version)
            // shape.Line.Color = Color.Blue;

            // Define output file path
            string outputPath = "ShapeWithLineStyle.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
