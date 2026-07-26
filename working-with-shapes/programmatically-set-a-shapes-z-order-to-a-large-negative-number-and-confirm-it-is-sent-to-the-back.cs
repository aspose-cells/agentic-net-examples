// Title: Send a shape to the back with a negative Z‑order in Aspose.Cells for .NET
// Description: Shows how to add a rectangle to a worksheet, move it to the back using Shape.ToFrontOrBack(-1), read its ZOrderPosition to verify the placement, and save the workbook.
// Keywords: Aspose.Cells | .NET | shape Z-order | ToFrontOrBack | send shape to back | ZOrderPosition | C# Excel shape ordering | negative Z order | move shape behind | Aspose.Cells shape layering
// Common Searches: Aspose.Cells move shape behind other objects | C# set shape ZOrderPosition negative | How to send shape to back in Excel using Aspose.Cells | Retrieve shape ZOrderPosition Aspose.Cells | Can ToFrontOrBack accept negative values
// Developer Intent: Programmatically place a shape behind all other worksheet objects by applying a negative Z‑order and confirm its position.
// Use Cases: Create a watermark that stays behind cell data. | Add background graphics without covering charts. | Ensure decorative shapes are hidden behind tables. | Prepare layered reports where annotations appear on top of base graphics.
// AI Prompts: Generate code to move all shapes in a worksheet to the back using a loop in Aspose.Cells for .NET. | Explain the range of ZOrderPosition values and how Aspose.Cells determines the backmost object. | Show how to validate that a shape is the lowest‑order object after calling ToFrontOrBack with a negative number.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Shows how to add a rectangle to a worksheet, move it to the back using Shape.ToFrontOrBack(-1), read its ZOrderPosition to verify the placement, and save the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

                // Send the shape to the back using a safe order value
                shape.ToFrontOrBack(-1); // move one position backward (to the back)

                // Verify the shape's Z-order position (lower values are farther back)
                Console.WriteLine("Shape ZOrderPosition after sending to back: " + shape.ZOrderPosition);

                // Save the workbook to verify the result visually if needed
                workbook.Save("ZOrderBackDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
