// Title: How to apply an outer shadow preset to shape text using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set outer shadow on shape text using TextOptions.Shadow | apply preset shadow type OffsetBottom to rectangle shape text in Aspose.Cells | configure shadow color, transparency, blur and distance for shape text in .NET | save workbook with shape text outer shadow Aspose.Cells example | C# tutorial for TextOptions.Shadow properties in Aspose.Cells
// Tags: Aspose.Cells shape text outer shadow preset | TextOptions.Shadow configuration C# | rectangle shape text shadow Aspose.Cells | shadow blur distance settings Aspose.Cells | save workbook with shape text effects

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace ApplyOuterShadowToShapeText
{
    // This C# example creates a new workbook, adds a rectangle shape, assigns text to the shape, and applies an outer shadow to the text by setting TextOptions.Shadow properties such as PresetType (OffsetBottom), gray color, 40% transparency, size 0.8, blur radius 20, and distance 6. The workbook is then saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upperRow, upperColumn, lowerRow, lowerColumn, width, height (in pixels)
                Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 200, 100);

                // Set the text of the shape
                shape.TextBody.Text = "Aspose.Cells Outer Shadow";

                // Apply an outer shadow to the shape's text
                shape.TextOptions.Shadow.PresetType = PresetShadowType.OffsetBottom; // outer shadow at bottom
                shape.TextOptions.Shadow.Color = workbook.CreateCellsColor();
                shape.TextOptions.Shadow.Color.Color = Color.Gray;
                shape.TextOptions.Shadow.Transparency = 0.4; // 40% transparent
                shape.TextOptions.Shadow.Size = 0.8;          // valid range: 0.0 – 2.0
                shape.TextOptions.Shadow.Blur = 20;          // blur radius
                shape.TextOptions.Shadow.Distance = 6;       // distance from the text

                // Save the workbook to an XLSX file
                string outputPath = "ShapeTextOuterShadow.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("The rectangle shape now has an outer shadow applied to its text.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
