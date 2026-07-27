// Title: Move a shape to the front layer (Z‑order) in Excel using Aspose.Cells for .NET
// Description: Creates a workbook, adds two overlapping rectangle shapes, and uses Shape.ToFrontOrBack(1) to bring the second shape to the front. An optional call with a negative value can send a shape to the back before saving the file.
// Keywords: Aspose.Cells | C# shape Z-order | Shape.ToFrontOrBack | Excel shape layering | move shape to front | send shape to back | Aspose.Cells drawing objects | overlapping shapes .NET
// Common Searches: Aspose.Cells move shape to front | C# change Z-order of Excel shapes | Shape.ToFrontOrBack example | bring rectangle to front Aspose.Cells | send shape to back Excel .NET
// Developer Intent: Adjust the Z‑order of a drawing object so that a specific shape appears above all other objects in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Ensure a title textbox stays visible above data bars in a generated report. | Place a company logo on top of decorative graphics in a dashboard sheet. | Prevent a chart from being hidden by a background shape in an automated spreadsheet.
// AI Prompts: Generate C# code that adds three overlapping shapes to a worksheet and moves the middle one to the front with Aspose.Cells. | Explain how Shape.ToFrontOrBack works, describing the effect of positive and negative order values. | Provide a C# snippet that iterates through all shapes in a worksheet and sends each one to the back in reverse order.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds two overlapping rectangle shapes, and uses Shape.ToFrontOrBack(1) to bring the second shape to the front. An optional call with a negative value can send a shape to the back before saving the file.
    public class ShapeToFrontDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two overlapping rectangle shapes
                Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
                Shape shape2 = worksheet.Shapes.AddRectangle(30, 30, 100, 100, 0, 0);

                // Bring shape2 to the front (positive order value)
                shape2.ToFrontOrBack(1);

                // Optionally, send shape1 to the back (negative order value)
                // shape1.ToFrontOrBack(-1);

                // Define output file path
                string outputPath = "ShapeToFrontDemo.xlsx";

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

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeToFrontDemo.Run();
        }
    }
}
