// Title: How to send a shape to the front layer (adjust Z‑order) in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move an Excel shape to the front of the Z‑order | How to bring overlapping shapes forward in an Excel file with Aspose.Cells | C# Aspose.Cells ToFrontOrBack method example for shape layering | Adjust shape layering in an .xlsx workbook using Aspose.Cells .NET
// Tags: Aspose.Cells ToFrontOrBack method | C# Excel shape Z-order manipulation | Aspose.Cells overlapping shape ordering | Excel .xlsx shape front layer API | Aspose.Cells shape layering technique

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeZOrderDemo
{
    // The example creates a new workbook, adds two overlapping rectangle shapes, and demonstrates how to change their Z‑order using the ToFrontOrBack method—first sending a shape backward and then bringing it to the front—before saving the workbook as ShapeZOrderDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping rectangle shapes
                // Shape 1 (bottom shape)
                // Parameters: topRow, top, leftColumn, left, height, width
                Shape shape1 = worksheet.Shapes.AddRectangle(
                    topRow: 5,
                    top: 10,
                    leftColumn: 2,
                    left: 10,
                    height: 120,
                    width: 200);

                shape1.Name = "BottomRectangle";
                shape1.Fill.FillType = FillType.Solid;
                shape1.Fill.SolidFill.Color = System.Drawing.Color.LightCoral;

                // Shape 2 (top shape, initially behind shape1)
                Shape shape2 = worksheet.Shapes.AddRectangle(
                    topRow: 6,
                    top: 30,
                    leftColumn: 3,
                    left: 30,
                    height: 120,
                    width: 200);

                shape2.Name = "TopRectangle";
                shape2.Fill.FillType = FillType.Solid;
                shape2.Fill.SolidFill.Color = System.Drawing.Color.LightSeaGreen;

                // Send shape2 to the back, then bring it to the front
                shape2.ToFrontOrBack(-1); // move backward
                shape2.ToFrontOrBack(2);  // move forward

                // Save the workbook
                string outputPath = "ShapeZOrderDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Shape 'TopRectangle' was sent to the back and then brought to the front using ToFrontOrBack.");
            }
            catch (Exception ex)
            {
                // Output any unexpected errors
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
