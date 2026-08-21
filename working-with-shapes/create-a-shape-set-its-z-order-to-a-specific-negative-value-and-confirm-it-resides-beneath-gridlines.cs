// Title: Aspose.Cells for .NET – Add a Rectangle Shape, Send It Behind Gridlines with Negative Z‑Order, and Verify Position
// Description: Demonstrates how to create a workbook, insert a rectangle shape, move the shape to the back of the Z‑order using a negative value so it sits beneath gridlines, read the read‑only ZOrderPosition to confirm placement, and save the file as an .xlsx document.
// Keywords: Aspose.Cells shape ZOrder | C# send shape behind gridlines | Aspose.Cells ToFrontOrBack negative | verify shape ZOrderPosition | add rectangle shape Aspose.Cells .NET | Excel shape layering Aspose
// Common Searches: Aspose.Cells move shape behind gridlines | C# set shape Z‑order negative Aspose.Cells | How to check ZOrderPosition of a shape in Aspose.Cells | Send shape to back of worksheet objects Aspose.Cells | Place rectangle behind cells using Aspose.Cells
// Developer Intent: Place a shape behind worksheet gridlines by applying a negative Z‑order value and confirm its position programmatically.
// Use Cases: Create a background watermark that does not obscure cell data. | Add decorative elements that stay under gridlines and content. | Ensure programmatically that shapes never cover important worksheet information.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle, sends it behind gridlines using ToFrontOrBack(-1), and logs the ZOrderPosition. | Explain the effect of negative values in the ToFrontOrBack method and how to validate the shape’s Z‑order. | Provide a sample that adds multiple shapes, orders them, and guarantees a specific shape remains behind all others and the gridlines.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Demonstrates how to create a workbook, insert a rectangle shape, move the shape to the back of the Z‑order using a negative value so it sits beneath gridlines, read the read‑only ZOrderPosition to confirm placement, and save the file as an .xlsx document.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top offset, left offset, height, width
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

                // Send the shape to the back of the Z-order (behind other objects, including gridlines)
                shape.ToFrontOrBack(-1);

                // Verify the Z-order position (read‑only property)
                Console.WriteLine("Shape ZOrderPosition: " + shape.ZOrderPosition);

                // Save the workbook to a file
                string outputPath = "ShapeBehindGridlines.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
