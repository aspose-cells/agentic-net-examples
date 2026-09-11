// Title: How to place a shape behind worksheet gridlines and confirm a negative Z‑Order using Aspose.Cells for .NET
// AI Prompts: Set the ZOrderPosition of a worksheet shape to a negative integer to position it behind the gridlines in an Aspose.Cells workbook. | Read the ZOrderPosition of the shape after setting it and raise an exception if the retrieved value is not negative.
// Common Searches: Aspose.Cells C# move shape behind gridlines | retrieve shape ZOrderPosition after setting negative value Aspose.Cells | verify shape Z-order is negative in .NET workbook | how to set shape ZOrderPosition to -1 using Aspose.Cells | place rectangle shape behind worksheet gridlines programmatically
// Tags: Aspose.Cells shape ZOrderPosition property | Aspose.Cells shape ZOrderPosition negative value | retrieve shape ZOrderPosition C# | verify shape Z-order Aspose.Cells | Aspose.Cells shape ZOrderPosition usage

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, assigns its ZOrderPosition a value of -1 to position it behind the worksheet gridlines, reads back the Z-order to ensure it is negative, throws an exception if the check fails, and saves the file as ShapeBehindGridlines.xlsx.
class ShapeZOrderExample
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Add a rectangle shape ----------
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // ---------- Move the shape behind gridlines ----------
            // Negative ZOrderPosition places the shape behind gridlines.
            shape.ZOrderPosition = -1;

            // ---------- Retrieve and verify the Z-order ----------
            int currentZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Current Z-Order: {currentZOrder}");

            // Ensure the Z-order is negative; otherwise raise an exception.
            if (currentZOrder >= 0)
            {
                throw new InvalidOperationException("Shape Z-Order is not negative as expected.");
            }

            // ---------- Save the workbook ----------
            string outputFile = "ShapeBehindGridlines.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
