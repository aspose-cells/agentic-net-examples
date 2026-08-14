// Title: Aspose.Cells .NET: Move a Shape to Front then Back and Verify Z‑Order
// Description: Demonstrates how to add overlapping rectangle shapes to a worksheet, bring one shape to the front with ToFrontOrBack(1), immediately send it to the back with ToFrontOrBack(0), and compare the initial and final ZOrderPosition values before saving the workbook.
// Keywords: Aspose.Cells shape layering | C# ToFrontOrBack method | ZOrderPosition Aspose.Cells | move shape to front .NET | send shape to back Excel | shape Z‑order verification | Aspose.Cells .NET examples
// Common Searches: Aspose.Cells move shape to front then back | How to change shape Z‑order in Aspose.Cells C# | ToFrontOrBack usage Aspose.Cells | Get shape ZOrderPosition after layering | C# example for shape front back Aspose.Cells
// Developer Intent: Programmatically adjust a shape’s stacking order—first to the front, then back—and confirm the resulting Z‑order positions.
// Use Cases: Temporarily highlight a shape for annotation and restore its original layer before exporting the workbook. | Validate that overlapping graphics retain the intended visual hierarchy in automated Excel report generation. | Debug shape layering issues by reading ZOrderPosition before and after layer changes.
// AI Prompts: Write C# code using Aspose.Cells to bring a shape to the front, then back, and display its ZOrderPosition values. | Explain the impact of ToFrontOrBack(1) and ToFrontOrBack(0) on shape Z‑order in Aspose.Cells. | Provide error‑handling best practices for moving shapes between front and back layers in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add overlapping rectangle shapes to a worksheet, bring one shape to the front with ToFrontOrBack(1), immediately send it to the back with ToFrontOrBack(0), and compare the initial and final ZOrderPosition values before saving the workbook.
public class ShapeFrontBackDemo
{
    public static void Main()
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
        Shape shape2 = worksheet.Shapes.AddRectangle(20, 20, 100, 100, 0, 0);

        // Record initial Z‑order positions
        int initialPos1 = shape1.ZOrderPosition;
        int initialPos2 = shape2.ZOrderPosition;

        try
        {
            // Bring shape2 to the front (parameter 1 = front)
            shape2.ToFrontOrBack(1);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error bringing shape to front: {ex.Message}");
        }

        try
        {
            // Send shape2 to the back (parameter 0 = back)
            shape2.ToFrontOrBack(0);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error sending shape to back: {ex.Message}");
        }

        // Record final Z‑order positions to verify the order
        int finalPos1 = shape1.ZOrderPosition;
        int finalPos2 = shape2.ZOrderPosition;

        Console.WriteLine($"Initial positions: shape1={initialPos1}, shape2={initialPos2}");
        Console.WriteLine($"Final positions:   shape1={finalPos1}, shape2={finalPos2}");

        // Save the workbook
        workbook.Save("ShapeFrontBackDemo.xlsx");
    }
}
