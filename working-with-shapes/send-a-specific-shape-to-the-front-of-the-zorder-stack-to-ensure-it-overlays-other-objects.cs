// Title: Move a Shape to the Front of the Z‑Order with Aspose.Cells .NET
// Description: C# sample that creates a workbook, adds overlapping rectangles, and uses Shape.ToFrontOrBack to raise one shape (or lower another) before saving the Excel file.
// Keywords: Aspose.Cells | C# shape Z order | ToFrontOrBack | Excel shape layering | move shape forward | send shape to back | Aspose.Cells .NET example
// Common Searches: Aspose.Cells bring shape to front | C# move Excel shape forward | ToFrontOrBack usage Aspose.Cells | change shape Z order in Excel .NET | layer shapes with Aspose.Cells
// Developer Intent: Reorder a specific shape so it appears above all other objects in an Excel worksheet.
// Use Cases: Overlay a label or annotation on a chart by moving it to the top layer. | Place a watermark image behind all content by sending it to the back of the stack. | Adjust shape order dynamically based on user input in a reporting application.
// AI Prompts: Generate C# code with Aspose.Cells that moves shape index 2 to the front of the Z‑order. | Show how to toggle a shape between front and back positions using ToFrontOrBack in Aspose.Cells. | Provide a script that orders multiple shapes so the largest rectangle ends up on top.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# sample that creates a workbook, adds overlapping rectangles, and uses Shape.ToFrontOrBack to raise one shape (or lower another) before saving the Excel file.
public class ShapeZOrderDemo
{
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
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add two overlapping rectangle shapes
        Shape shape1 = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
        Shape shape2 = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

        // Bring shape2 to the front of the Z‑order stack
        shape2.ToFrontOrBack(1); // positive value moves shape forward

        // (Optional) Send shape1 to the back of the Z‑order stack
        shape1.ToFrontOrBack(-1); // negative value moves shape backward

        // Save the workbook to a file
        string outputPath = "ShapeZOrderDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
