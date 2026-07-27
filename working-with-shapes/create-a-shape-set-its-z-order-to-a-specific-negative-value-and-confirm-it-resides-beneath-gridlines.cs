// Title: Aspose.Cells .NET: Add Rectangle Shape, Send Behind Gridlines via Negative Z‑Order, Verify Placement
// Description: Demonstrates how to create a workbook, insert a rectangle shape, shift it one step back in the Z‑order using ToFrontOrBack(-1) so it sits beneath the worksheet gridlines, output its ZOrderPosition, and save the file.
// Keywords: Aspose.Cells shape Z-order | ToFrontOrBack negative value | shape behind gridlines | ZOrderPosition verification | Aspose.Cells C# shape layering
// Common Searches: Aspose.Cells send shape behind gridlines | How to use ToFrontOrBack in Aspose.Cells | Retrieve ZOrderPosition of a shape .NET | Move shape backward in worksheet Aspose.Cells
// Developer Intent: Place a rectangle shape under the worksheet gridlines by adjusting its Z‑order and confirm the ordering programmatically.
// Use Cases: Add background graphics that do not obscure cell data. | Create watermarks that appear beneath gridlines for reports. | Control visual hierarchy when multiple shapes overlap on a sheet.
// AI Prompts: Write C# code with Aspose.Cells that adds a shape, moves it two positions back in the Z‑order, and prints the new ZOrderPosition. | Explain the behavior of ToFrontOrBack for worksheet shapes and how ZOrderPosition reflects their stacking order. | Show an alternative method to send a shape to the back of all objects in an Aspose.Cells worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, insert a rectangle shape, shift it one step back in the Z‑order using ToFrontOrBack(-1) so it sits beneath the worksheet gridlines, output its ZOrderPosition, and save the file.
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
            // Parameters: upper left row, upper left column, width, height, upper left pixel offset X, upper left pixel offset Y
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);

            // Move the shape one position backward in the Z-order (negative value)
            // Use -1 to avoid exceeding collection bounds
            shape.ToFrontOrBack(-1);

            // Output the Z-order position to confirm its current order
            Console.WriteLine("Shape ZOrderPosition: " + shape.ZOrderPosition);

            // Save the workbook
            workbook.Save("ShapeBehindGridlines.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
