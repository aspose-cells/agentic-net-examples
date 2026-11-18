// Title: Create a rectangle shape with dashed line style and 2.5‑point thickness using Aspose.Cells for .NET
// Common Searches: how to add a rectangle shape with dashed border in Aspose.Cells C# | setting line dash style and weight for shapes in Aspose.Cells .NET | Aspose.Cells C# shape line formatting example | apply custom line thickness to a shape in an Excel workbook using Aspose.Cells | define dash pattern for rectangle shape using Aspose.Cells API
// Tags: shape dash style configuration Aspose.Cells | custom line thickness for Excel shape C# | MsoLineDashStyle dash usage Aspose.Cells | apply border formatting to rectangle shape .NET | Aspose.Cells shape line formatting API

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a rectangle shape on the first worksheet, sets its line dash style to dash and its weight to 2.5 points, and saves the result as ShapeLineStyle.xlsx.
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
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 50, 50, 100, 200);

            // Define the line style for the shape
            shape.Line.DashStyle = MsoLineDashStyle.Dash;   // dash pattern
            shape.Line.Weight = 2.5f;                      // line thickness in points

            // Optional: set line color (commented out because LineFormat.Color may not be available in some versions)
            // shape.Line.Color = Color.Blue;

            // Save the workbook to a file
            string outputPath = "ShapeLineStyle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
