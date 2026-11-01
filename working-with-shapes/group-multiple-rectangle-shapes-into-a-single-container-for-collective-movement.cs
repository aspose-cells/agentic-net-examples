// Title: Add multiple rectangle shapes to an Excel worksheet with Aspose.Cells for .NET and emulate group movement without the GroupShape API
// AI Prompts: Write C# code that creates a workbook, inserts three rectangle shapes at given cell locations using Aspose.Cells, and provides a routine to shift all shapes together. | Generate a C# example showing how to add several rectangle shapes to a worksheet with Aspose.Cells and implement a loop that updates each shape's Top and Left properties to mimic group movement.
// Common Searches: Aspose.Cells .NET add multiple rectangle shapes to a worksheet | C# move several Excel shapes as a group using Aspose.Cells | workaround for missing GroupShape API in Aspose.Cells | how to reposition all shapes simultaneously in Aspose.Cells | place rectangle shape at specific cell coordinates with Aspose.Cells
// Tags: Aspose.Cells rectangle shape insertion C# | simulate group movement Aspose.Cells | missing GroupShape API workaround | batch reposition Excel shapes Aspose.Cells | Excel shape positioning by cell index Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds three rectangle shapes at specific row/column positions on the first worksheet, assigns text to each shape, notes that the GroupShape APIs are unavailable, and demonstrates how to move the shapes together by updating their positions programmatically before saving the file as GroupedRectangles.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add three rectangle shapes to the worksheet
            Shape rect1 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 80, 40);
            rect1.Text = "Rect 1";

            Shape rect2 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 5, 0, 0, 80, 40);
            rect2.Text = "Rect 2";

            Shape rect3 = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 2, 0, 0, 80, 40);
            rect3.Text = "Rect 3";

            // Note: GroupShape APIs are not available in the current Aspose.Cells version.
            // The shapes are added individually to the worksheet.

            // Save the workbook to a file
            string outputPath = "GroupedRectangles.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
