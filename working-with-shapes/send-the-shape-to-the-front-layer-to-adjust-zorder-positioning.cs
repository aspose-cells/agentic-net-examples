// Title: Send a shape to the front layer (Z‑order) with Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds two overlapping rectangle shapes, uses the ToFrontOrBack method with a positive value to move a shape to the front (and optionally a negative value to send a shape to the back), and saves the result as ShapeZOrderDemo.xlsx.
// Keywords: Aspose.Cells | C# | shape Z-order | ToFrontOrBack | bring shape to front | send shape to back | Excel shape layering | .NET workbook | overlapping shapes
// Common Searches: Aspose.Cells change shape Z order | C# bring shape to front in Excel | send shape to back Aspose.Cells | adjust shape layering programmatically | move overlapping shapes in workbook
// Developer Intent: Reorder a specific shape’s Z‑order so it appears on the front (or back) layer of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Ensure a chart legend or annotation stays visible above other overlapping objects. | Create complex diagrams where certain shapes must be foreground elements. | Programmatically prepare reports with controlled visual hierarchy before distribution.
// AI Prompts: Generate C# code with Aspose.Cells that moves the third shape in a worksheet to the front while leaving other shapes unchanged. | Explain the numeric parameters of the ToFrontOrBack method and how they affect shape positioning. | Write a script that iterates through all worksheet shapes and sends any shape named "Background" to the back layer.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds two overlapping rectangle shapes, uses the ToFrontOrBack method with a positive value to move a shape to the front (and optionally a negative value to send a shape to the back), and saves the result as ShapeZOrderDemo.xlsx.
    public class ShapeToFrontDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two overlapping shapes
            Shape shape1 = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(20, 20, 100, 100, 0, 0);

            // Bring shape2 to the front (positive value)
            shape2.ToFrontOrBack(1);

            // Send shape1 to the back (negative value) – optional
            shape1.ToFrontOrBack(-1);

            // Save the workbook
            workbook.Save("ShapeZOrderDemo.xlsx");
        }

        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
