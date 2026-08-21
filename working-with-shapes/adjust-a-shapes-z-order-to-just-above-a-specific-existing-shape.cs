// Title: C# – Move a shape just above another shape by adjusting ZOrderPosition in Aspose.Cells
// Description: Demonstrates how to create a workbook, add two rectangle shapes, and set the ZOrderPosition of one shape to be one level higher than a reference shape so it renders directly above it, then saves the file.
// Keywords: Aspose.Cells C# shape ZOrderPosition | adjust shape stacking order .NET | bring shape to front Aspose.Cells | Excel shape layering programmatically | change shape order Aspose.Cells
// Common Searches: Aspose.Cells set shape above another shape | C# change ZOrderPosition of Excel shape | how to bring a shape forward in Aspose.Cells | move shape in front of another shape .NET | shape layering example Aspose.Cells
// Developer Intent: Set shapeA's Z-order so it appears directly in front of shapeB.
// Use Cases: Overlay a label on a chart to show dynamic titles. | Ensure comment or annotation shapes stay visible above data markers. | Display a custom tooltip shape above a selected cell range.
// AI Prompts: Generate C# code that moves a specific shape above another without hard‑coding ZOrderPosition values. | Show how to programmatically bring a chosen shape to the front of all worksheet shapes using Aspose.Cells. | Explain how to swap the ZOrderPosition of two shapes in a .NET workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderExample
{
    // Demonstrates how to create a workbook, add two rectangle shapes, and set the ZOrderPosition of one shape to be one level higher than a reference shape so it renders directly above it, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two rectangle shapes
            // shapeA will be the shape we want to move
            Shape shapeA = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);
            // shapeB is the reference shape
            Shape shapeB = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0);

            // Initial Z-order (optional, just for clarity)
            // Lower ZOrderPosition means closer to the back
            shapeA.ZOrderPosition = 0; // back
            shapeB.ZOrderPosition = 1; // front of shapeA

            // Adjust shapeA to be just above shapeB
            // Set shapeA's ZOrderPosition to shapeB's position + 1
            shapeA.ZOrderPosition = shapeB.ZOrderPosition + 1;

            // Save the workbook to verify the result
            workbook.Save("ZOrderAdjusted.xlsx");
        }
    }
}
