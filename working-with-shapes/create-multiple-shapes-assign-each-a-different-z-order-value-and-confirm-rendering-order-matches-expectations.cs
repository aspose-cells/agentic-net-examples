// Title: Assign Z-order to multiple rectangle shapes in an Excel worksheet and verify rendering order with Aspose.Cells for .NET
// AI Prompts: Create three rectangle shapes on a worksheet, set each shape's ZOrderPosition (top, middle, bottom), and output the shape names sorted by their Z-order. | Generate an Excel file where shapes are layered using ZOrderPosition, then iterate through the Shapes collection ordered by ZOrderPosition to display the rendering sequence in the console. | Use Aspose.Cells to add rectangle shapes, assign custom Z-order indices, sort the shapes by ZOrderPosition, and save the workbook.
// Common Searches: Aspose.Cells how to change shape layering order in C# | C# set ZOrderPosition for multiple shapes in Excel workbook | retrieve shapes sorted by ZOrderPosition using Aspose.Cells | verify shape rendering sequence in generated Excel file Aspose.Cells .NET | example of assigning Z-order to rectangle shapes with Aspose.Cells
// Tags: Aspose.Cells shape Z-order assignment C# | Excel rectangle shape layering Aspose.Cells | C# add rectangle shapes Aspose.Cells | retrieve shapes by ZOrderPosition Aspose.Cells | save workbook with layered shapes Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, adds three rectangle shapes to the first worksheet, assigns each a distinct ZOrderPosition (2 for topmost, 0 for bottom, 1 for middle), sorts the shapes by this property to display the rendering order from bottom to top, and saves the file as ShapesZOrder.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add three rectangle shapes with different Z-order values
            Shape shape1 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 0, 0, 0, 100, 50);
            shape1.Name = "Shape1";
            shape1.ZOrderPosition = 2; // Topmost

            Shape shape2 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 0, 0, 0, 100, 50);
            shape2.Name = "Shape2";
            shape2.ZOrderPosition = 0; // Bottom

            Shape shape3 = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 3, 0, 0, 0, 100, 50);
            shape3.Name = "Shape3";
            shape3.ZOrderPosition = 1; // Middle

            // Verify rendering order by sorting shapes based on ZOrderPosition (bottom to top)
            var orderedShapes = sheet.Shapes.Cast<Shape>()
                                            .OrderBy(s => s.ZOrderPosition)
                                            .ToList();

            Console.WriteLine("Shapes rendering order (bottom to top):");
            foreach (var s in orderedShapes)
            {
                Console.WriteLine($"{s.Name} - ZOrderPosition: {s.ZOrderPosition}");
            }

            // Save the workbook to a file
            string outputPath = "ShapesZOrder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
