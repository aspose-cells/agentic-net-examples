// Title: Programmatically bring a shape to the front, send it to the back, and compare Z-order positions in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a rectangle shape to a worksheet, reads its initial ZOrderPosition, sets the shape to the front by assigning the highest Z-order index, then moves it to the back by setting ZOrderPosition to 0, and prints the before/after values using Aspose.Cells. | Show how to retrieve a shape's Z-order index, modify it to bring the shape forward or send it backward, and verify the changes in an Excel workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set shape ZOrderPosition to bring shape to front | How to move an Excel shape to the back using Aspose.Cells .NET | Retrieve and compare shape Z-order index in Aspose.Cells workbook | C# example for changing Z-order of shapes in Excel with Aspose.Cells | Aspose.Cells shape ordering front back Z-order manipulation
// Tags: Aspose.Cells shape ZOrderPosition | C# bring Excel shape to front Aspose.Cells | C# send Excel shape to back Aspose.Cells | compare shape Z-order Aspose.Cells | rectangle shape Z-order Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a rectangle shape, captures its initial ZOrderPosition, moves the shape to the front by setting the Z-order to the highest index, then sends it to the back by setting ZOrderPosition to zero, compares the indices, and saves the workbook.
class ShapeZOrderDemo
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
            // Parameters: shape type, upper left row, upper left column, upper left row offset, upper left column offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 100, 100);

            // Capture the initial Z-order position
            int initialZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Initial Z-order: {initialZOrder}");

            // Bring the shape to the front by setting its Z-order to the maximum index
            shape.ZOrderPosition = sheet.Shapes.Count - 1;
            int frontZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Z-order after bringing to front: {frontZOrder}");

            // Send the shape to the back by setting its Z-order to zero
            shape.ZOrderPosition = 0;
            int backZOrder = shape.ZOrderPosition;
            Console.WriteLine($"Z-order after sending to back: {backZOrder}");

            // Verify the Z-order manipulation
            if (initialZOrder == backZOrder && frontZOrder > initialZOrder)
            {
                Console.WriteLine("Z-order manipulation succeeded.");
            }
            else
            {
                Console.WriteLine("Z-order manipulation did not behave as expected.");
            }

            // Save the workbook (optional, just to demonstrate lifecycle usage)
            string outputPath = "ShapeZOrderDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
