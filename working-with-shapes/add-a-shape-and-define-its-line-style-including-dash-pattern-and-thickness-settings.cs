// Title: Create a rectangle shape with a dashed 2‑point line in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that inserts a rectangle shape into the first worksheet of a new workbook and applies a dashed line with a 2‑point weight using Aspose.Cells. | Show how to configure the line dash style and line thickness of a shape in an Excel file with Aspose.Cells for .NET. | Generate an example that creates an Excel workbook, adds a rectangle shape, sets its line to Dash and thickness to 2 points, then saves the file as .xlsx.
// Common Searches: Aspose.Cells how to add a shape with custom line style in C# | set dash pattern and line thickness for Excel shape using Aspose.Cells .NET | C# example for drawing a rectangle with dashed border in an Excel workbook with Aspose.Cells | Aspose.Cells shape line formatting options C# | apply line weight to rectangle shape in Aspose.Cells workbook
// Tags: Aspose.Cells add rectangle shape | Aspose.Cells shape line dash style | Aspose.Cells set line weight | Aspose.Cells C# shape formatting | Aspose.Cells export to xlsx with styled shape

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// // Creates a new workbook, adds a rectangle shape to the first worksheet, sets its line dash style to Dash and weight to 2 points, and saves the file as ShapeWithLineStyle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, row offset, column offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // row offset (in points)
                0,   // column offset (in points)
                100, // height (in points)
                200  // width (in points)
            );

            // Define the line style
            shape.Line.DashStyle = MsoLineDashStyle.Dash; // dash pattern
            shape.Line.Weight = 2.0;                     // line thickness in points
            // Optional line color – comment out if not supported by the current Aspose.Cells version
            // shape.Line.Color = Color.Blue;

            // Save the workbook (lifecycle save)
            string outputPath = "ShapeWithLineStyle.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
