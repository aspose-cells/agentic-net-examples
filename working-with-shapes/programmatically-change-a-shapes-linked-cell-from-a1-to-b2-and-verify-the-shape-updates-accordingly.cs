// Title: How to change a shape's linked cell from A1 to B2 and update its text using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change shape linked cell from A1 to B2 | update rectangle shape text after modifying LinkedCell property in Aspose.Cells | programmatically set shape.LinkedCell to another cell with Aspose.Cells .NET | link shape to a different worksheet cell using Aspose.Cells API | verify shape reflects new linked cell value in an Aspose.Cells workbook
// Tags: Aspose.Cells shape LinkedCell update | C# rectangle shape linked cell change | synchronizing shape text with worksheet cell Aspose.Cells | programmatic shape linking to cell Aspose.Cells | Aspose.Cells workbook shape text refresh

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape linked to cell A1, sets the shape's text from the cell value, then changes the shape's LinkedCell to B2, updates the new cell's value, refreshes the shape text, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape and link it to cell A1
            // Parameters: drawing type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row
                0,    // upper left column
                0,    // top (in pixels)
                0,    // left (in pixels)
                100,  // height (in pixels)
                100   // width (in pixels)
            );

            shape.Placement = PlacementType.FreeFloating; // allow linked cell behavior
            shape.LinkedCell = "A1";

            // Set initial value in A1 and update shape text manually
            sheet.Cells["A1"].PutValue("Old Value");
            shape.Text = sheet.Cells["A1"].StringValue; // reflect linked cell value

            Console.WriteLine("After linking to A1, shape text: " + shape.Text); // Expected: Old Value

            // Change the linked cell to B2
            shape.LinkedCell = "B2";

            // Set value in B2 and update shape text
            sheet.Cells["B2"].PutValue("New Value");
            shape.Text = sheet.Cells["B2"].StringValue; // reflect new linked cell value

            Console.WriteLine("After linking to B2, shape text: " + shape.Text); // Expected: New Value

            // Save the workbook (optional, just to demonstrate lifecycle)
            string outputPath = Path.Combine(Environment.CurrentDirectory, "LinkedShapeDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
