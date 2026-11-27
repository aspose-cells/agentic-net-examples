// Title: How to rotate an auto shape 90° to a diagonal position using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set shape rotation angle to 90 degrees | rotate rectangle auto shape diagonally in an Excel workbook with Aspose.Cells | Aspose.Cells .NET example for rotating shapes programmatically | how to apply rotation to a shape in an Aspose.Cells worksheet using C# | set diagonal orientation for a rectangle shape in Excel using Aspose.Cells
// Tags: Aspose.Cells shape rotation | C# auto shape rotation | set rotation angle Aspose.Cells | diagonal shape orientation .NET | Excel workbook shape rotation example

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeRotationDemo
{
    // The sample creates a new workbook, adds a rectangle auto shape to the first worksheet, assigns text, sets its RotationAngle property to 90 degrees for a diagonal layout, and saves the file as ShapeRotationDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle auto shape to the worksheet
                // Parameters: AutoShapeType, upperLeftRow, upperLeftColumn, top, left, height, width
                Shape shape = worksheet.Shapes.AddAutoShape(
                    AutoShapeType.Rectangle, // shape type
                    2,    // upper left row (zero‑based)
                    2,    // upper left column (zero‑based)
                    5,    // top offset in pixels
                    5,    // left offset in pixels
                    100,  // height in pixels
                    200   // width in pixels
                );

                // Optional: give the shape some text so we can see it after rotation
                shape.Text = "Diagonal Shape";

                // Set the rotation angle to 90 degrees (clockwise)
                shape.RotationAngle = 90f;

                // Save the workbook to a file
                string outputPath = "ShapeRotationDemo.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The shape has been rotated 90 degrees for diagonal orientation.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook or setting the shape:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
