// Title: How to set a shape's Z‑order to a negative value to send it to the back with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to negative | move Excel shape behind other objects using Aspose.Cells | retrieve current Z‑order of a shape after setting in Aspose.Cells | example of shape layering order manipulation in Aspose.Cells .NET | send rectangle shape to back of worksheet programmatically Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition negative | C# Excel shape layering Aspose.Cells | set shape Z-order Aspose.Cells .NET | move shape to back worksheet Aspose.Cells | verify shape Z-order value C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape to the first worksheet, assigns a large negative ZOrderPosition to push the shape to the back, prints the resulting Z‑order value, and saves the workbook as ZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Set the Z‑order to a large negative number to push the shape to the back
            shape.ZOrderPosition = -1000;

            // Confirm the Z‑order value (the shape should now be at the back)
            int currentZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Shape Z‑order after setting: {currentZOrder}");

            // Save the workbook
            workbook.Save("ZOrderDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
