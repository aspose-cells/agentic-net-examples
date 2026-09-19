// Title: How to send a worksheet shape to the back by setting a large negative ZOrderPosition with Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a rectangle shape to the first worksheet, assign ZOrderPosition = -10000, and save the workbook as ShapeZOrderDemo.xlsx. | After setting a negative ZOrderPosition, retrieve and print the shape's ZOrderPosition to confirm it is behind other drawings. | Wrap the shape Z-order adjustment in a try‑catch block and log any exceptions that occur during workbook creation or saving.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition negative value to move shape behind others | send worksheet drawing to back using ZOrderPosition in Aspose.Cells .NET | programmatically change drawing order of shapes in an Excel file with Aspose.Cells
// Tags: Aspose.Cells shape ZOrderPosition example | C# move worksheet shape to back | Aspose.Cells drawing order manipulation | set negative ZOrderPosition .NET | Excel shape layering Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a new workbook, adds a rectangle shape to the first worksheet, sets its ZOrderPosition to -10000 to place it behind all other objects, prints the Z-order value for verification, and saves the workbook as ShapeZOrderDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, height, width
            Shape rect = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 50);

            // Set the Z-order to a large negative number to send it to the back
            // In Aspose.Cells, lower Z-order values are rendered behind higher values
            rect.ZOrderPosition = -10000;

            // Confirm the Z-order value
            Console.WriteLine($"Shape ZOrderPosition set to: {rect.ZOrderPosition}");
            // Since we only have one shape, it is now at the back of the drawing order.

            // Save the workbook to a file (lifecycle rule: save)
            string outputPath = "ShapeZOrderDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
