// Title: C# – Move a Shape Behind Gridlines and Retrieve Its Z‑Order with Aspose.Cells
// Description: Demonstrates how to add a rectangle to a worksheet, send it to the back of all objects and gridlines using ToFrontOrBack(1), read the resulting ZOrderPosition, and save the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells shape back | C# send shape behind gridlines | Aspose.Cells ZOrderPosition | shape layering Excel .NET | ToFrontOrBack Aspose.Cells | C# Excel shape ordering | Aspose.Cells workbook save | global .NET Excel library
// Common Searches: Aspose.Cells move shape to back | Get Z‑order of a shape in C# Excel | Place shape behind gridlines Aspose | How to use ToFrontOrBack in Aspose.Cells | C# retrieve shape ZOrderPosition
// Developer Intent: Place a worksheet shape behind all other objects and gridlines, then obtain its Z‑order index to confirm the ordering.
// Use Cases: Create a background watermark that stays under gridlines. | Programmatically reorder multiple shapes so a specific one becomes the bottom layer. | Validate shape ordering before exporting the workbook to Excel.
// AI Prompts: Write C# code that adds a shape, sends it behind gridlines with Aspose.Cells, and checks that its ZOrderPosition reflects the lowest layer. | Explain the effect of ToFrontOrBack(1) on shape layering and how ZOrderPosition is calculated in Aspose.Cells for .NET. | Provide a script to reorder several shapes so that a chosen shape ends up behind all others and gridlines.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle to a worksheet, send it to the back of all objects and gridlines using ToFrontOrBack(1), read the resulting ZOrderPosition, and save the workbook as an XLSX file using Aspose.Cells for .NET.
class MoveShapeBehindGridlines
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (row, column, upper left row offset, upper left column offset, width, height)
            Shape shape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 100);

            // Send the shape to the back (behind other objects and gridlines)
            // 0 = bring to front, 1 = send to back
            shape.ToFrontOrBack(1);

            // Retrieve the Z-order position (non‑negative index)
            int zPos = shape.ZOrderPosition;
            Console.WriteLine("Z-order position: " + zPos);
            Console.WriteLine("Shape has been sent to the back (behind gridlines).");

            // Save the workbook
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
