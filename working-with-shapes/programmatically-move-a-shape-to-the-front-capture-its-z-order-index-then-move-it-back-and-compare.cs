// Title: Aspose.Cells for .NET – Move a Shape to Front, Capture Z‑Order, Return to Back
// Description: Demonstrates how to add overlapping shapes to a worksheet, read their ZOrderPosition, bring a shape to the front with ToFrontOrBack(1), send it to the back with ToFrontOrBack(-1), and compare the indices before saving the workbook.
// Keywords: Aspose.Cells .NET shape ZOrderPosition | ToFrontOrBack method | move shape to front | move shape to back | shape layering Excel | Aspose.Cells shape ordering
// Common Searches: Aspose.Cells change shape Z‑order | Get shape ZOrderPosition .NET | Bring shape to front Aspose.Cells | Send shape to back Aspose.Cells | Compare shape Z‑order after ToFrontOrBack
// Developer Intent: Learn how to programmatically adjust a shape's Z‑order, retrieve its index, and verify the change using Aspose.Cells for .NET.
// Use Cases: Ensure critical graphics appear on top in generated Excel reports. | Debug shape layering by logging Z‑order values during workbook creation. | Automate dynamic re‑ordering of overlapping shapes based on business rules.
// AI Prompts: Show C# code that moves a shape to the front, reads its ZOrderPosition, then moves it to the back and validates the order with Aspose.Cells. | Generate a script that lists all shapes in a worksheet and prints each shape's ZOrderPosition. | Explain how ToFrontOrBack(1) and ToFrontOrBack(-1) affect shape layering in an Excel file created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add overlapping shapes to a worksheet, read their ZOrderPosition, bring a shape to the front with ToFrontOrBack(1), send it to the back with ToFrontOrBack(-1), and compare the indices before saving the workbook.
class ShapeZOrderDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two overlapping rectangle shapes
        Shape shape1 = sheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
        Shape shape2 = sheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

        // Capture initial Z-order positions
        int initialPos1 = shape1.ZOrderPosition;
        int initialPos2 = shape2.ZOrderPosition;

        // Bring shape2 to the front (positive order)
        shape2.ToFrontOrBack(1);
        int frontPos = shape2.ZOrderPosition;

        // Send shape2 to the back (negative order)
        shape2.ToFrontOrBack(-1);
        int backPos = shape2.ZOrderPosition;

        // Output the captured positions and comparison result
        Console.WriteLine($"Initial ZOrder: shape1 = {initialPos1}, shape2 = {initialPos2}");
        Console.WriteLine($"After ToFrontOrBack(1): shape2 ZOrder = {frontPos}");
        Console.WriteLine($"After ToFrontOrBack(-1): shape2 ZOrder = {backPos}");
        Console.WriteLine($"Front position > Back position? {frontPos > backPos}");

        // Save the workbook to verify the shapes are present
        workbook.Save("ShapeZOrderDemo.xlsx");
    }
}
