// Title: C# – Change Shape Z‑Order (Bring to Front / Send to Back) in Excel with Aspose.Cells
// Description: Shows how to add overlapping rectangle shapes to a worksheet and adjust their Z‑order using Aspose.Cells for .NET. The sample employs Shape.ToFrontOrBack to move a shape forward or backward before saving the workbook.
// Keywords: Aspose.Cells | C# shape Z-order | ToFrontOrBack | bring shape to front | send shape to back | overlapping shapes Excel | modify shape order | Excel worksheet shapes | .NET Aspose.Cells example | shape layering
// Common Searches: Aspose.Cells move shape to front | C# set Z-order of Excel shapes | how to bring a rectangle shape forward in Aspose.Cells | send shape backward programmatically Aspose.Cells | Shape.ToFrontOrBack usage | change layering of shapes in Excel using .NET
// Developer Intent: Reorder overlapping shapes so a chosen shape appears above or below the others in an Excel file.
// Use Cases: Create overlapping charts or images and control which element is visible. | Implement a drag‑and‑drop UI where users can change the stacking order of shapes. | Generate reports with layered graphics where the foreground shape must be highlighted. | Programmatically adjust Z‑order in a loop for dynamic dashboards or visualizations.
// AI Prompts: Write C# code that iterates through all worksheet shapes and moves each one to the front using ToFrontOrBack(1). | Provide a WinForms button‑click handler that toggles two shapes between front and back positions with Aspose.Cells. | Explain how positive, zero, and negative values affect Shape.ToFrontOrBack and how to place a shape at the absolute front or back. | Generate a PowerShell script that reorders shapes in an Excel workbook based on a priority list using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Shows how to add overlapping rectangle shapes to a worksheet and adjust their Z‑order using Aspose.Cells for .NET. The sample employs Shape.ToFrontOrBack to move a shape forward or backward before saving the workbook.
    public class ShapeZOrderDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two overlapping rectangle shapes
            Shape shape1 = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);
            Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

            // Bring shape2 to the front (positive value moves forward)
            shape2.ToFrontOrBack(1);

            // Send shape1 to the back (negative value moves backward)
            shape1.ToFrontOrBack(-1);

            // Save the workbook
            workbook.Save("ShapeZOrderDemo.xlsx");
            Console.WriteLine("Workbook saved as ShapeZOrderDemo.xlsx");
        }
    }
}
