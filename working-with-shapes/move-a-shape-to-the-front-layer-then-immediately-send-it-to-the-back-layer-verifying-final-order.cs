// Title: Move a rectangle shape to the front layer, then send it to the back layer and verify Z-order using Aspose.Cells for .NET
// AI Prompts: Create a C# program with Aspose.Cells that adds a rectangle shape, sets its ZOrderPosition to the highest index, then resets it to zero, and prints whether the shape is now the backmost item. | Write code that manipulates a worksheet shape's Z-order by assigning shape.ZOrderPosition = sheet.Shapes.Count-1 followed by shape.ZOrderPosition = 0, then confirms the final order in the Shapes collection.
// Common Searches: Aspose.Cells C# change shape order front to back example | How to set ZOrderPosition for a shape in Aspose.Cells | Check if a shape is backmost after SendToBack in Aspose.Cells | C# Aspose.Cells move rectangle shape behind other objects
// Tags: Aspose.Cells shape Z-order adjustment | C# bring shape to front Aspose.Cells | C# send shape to back Aspose.Cells | Aspose.Cells verify shape backmost order

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangle shape, moves it to the front by setting the highest ZOrderPosition, then sends it to the back by resetting the position to zero, verifies that the shape becomes the first item in the worksheet's Shapes collection, and saves the file as ShapeZOrder.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            var shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);

            // Bring the shape to the front layer by setting the highest Z-order position
            shape.ZOrderPosition = sheet.Shapes.Count - 1;

            // Send the shape to the back layer by setting the lowest Z-order position
            shape.ZOrderPosition = 0;

            // Verify final order: after SendToBack the shape should be the first (backmost) in the collection
            bool isBackmost = sheet.Shapes[0] == shape;
            Console.WriteLine("Shape is backmost: " + isBackmost);

            // Save the workbook
            workbook.Save("ShapeZOrder.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
