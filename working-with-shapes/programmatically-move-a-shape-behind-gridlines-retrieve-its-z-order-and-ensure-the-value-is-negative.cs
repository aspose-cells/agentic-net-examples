// Title: Aspose.Cells for .NET: Move a Shape Behind Gridlines and Retrieve Its Z‑Order Position
// Description: Demonstrates how to add a rectangle shape to a worksheet, send it behind the gridlines using ToFrontOrBack(1), read the shape's ZOrderPosition (which becomes the lowest index), and save the workbook. Shows proper layering of shapes under cells with Aspose.Cells.
// Keywords: Aspose.Cells shape behind gridlines | C# ToFrontOrBack Aspose.Cells | ZOrderPosition shape Aspose.Cells | move shape to back worksheet | shape layering Aspose.Cells .NET
// Common Searches: how to send a shape to the back of gridlines using Aspose.Cells C# | retrieve ZOrderPosition after moving shape behind cells | Aspose.Cells shape layering behind worksheet gridlines | C# Aspose.Cells move shape behind cells example | what does ZOrderPosition value mean in Aspose.Cells
// Developer Intent: Place a shape behind worksheet gridlines, read its Z‑order index, and confirm it is the lowest order.
// Use Cases: Add a watermark that stays under the data grid for printable reports. | Create background graphics that should not obscure cell content. | Programmatically adjust shape layering to ensure annotations appear beneath worksheet elements.
// AI Prompts: Generate C# code with Aspose.Cells that adds a shape, moves it behind gridlines, and returns the ZOrderPosition. | Explain the effect of ToFrontOrBack(1) on a shape's ZOrderPosition in Aspose.Cells. | Write a unit test in C# that asserts the shape's ZOrderPosition is the minimum value after calling ToFrontOrBack(1).

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeZOrderDemo
{
    // Demonstrates how to add a rectangle shape to a worksheet, send it behind the gridlines using ToFrontOrBack(1), read the shape's ZOrderPosition (which becomes the lowest index), and save the workbook. Shows proper layering of shapes under cells with Aspose.Cells.
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
                // Parameters: upper left row, upper left column, upper left pixel offset Y, upper left pixel offset X, height, width
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

                // Send the shape behind the gridlines (to the back)
                // The method expects 1 to move the shape to the back
                shape.ToFrontOrBack(1);

                // Retrieve and display the current Z-order position
                int zOrder = shape.ZOrderPosition;
                Console.WriteLine("ZOrderPosition after sending to back: " + zOrder);

                // Save the workbook to a file
                workbook.Save("ShapeBehindGridlines.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
