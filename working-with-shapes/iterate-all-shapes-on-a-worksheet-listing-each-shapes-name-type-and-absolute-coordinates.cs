// Title: Enumerate Worksheet Shapes and Retrieve Name, Type, and Absolute Coordinates with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds a rectangle, oval, and textbox, then iterates the worksheet's Shapes collection. For each shape it extracts the Name, MsoDrawingType, Left, Top, Width, and Height properties and writes the information to the console before optionally saving the file.
// Keywords: Aspose.Cells | C# | .NET | shape enumeration | worksheet shapes | shape coordinates | MsoDrawingType | Excel drawing objects | list shape properties | shape left top width height
// Common Searches: Aspose.Cells iterate shapes C# | how to list shape name and type in Excel using Aspose | retrieve shape coordinates with Aspose.Cells for .NET | enumerate all drawing objects in a worksheet Aspose | get left top width height of shapes in Aspose.Cells
// Developer Intent: Extract every shape’s identifier, drawing type, and pixel position/size from a worksheet.
// Use Cases: Generate an inventory of all drawing objects for documentation or compliance audits. | Export shape layout data to external systems that require exact positioning information. | Validate shape placement against design guidelines before producing printable reports.
// AI Prompts: Write C# code using Aspose.Cells to loop through all worksheet shapes and print each shape's Name, MsoDrawingType, Left, Top, Width, and Height. | Show how to collect shape metadata into a DataTable or List for further processing with Aspose.Cells for .NET. | Provide an example that saves the enumerated shape details to a CSV file after iterating the Shapes collection.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeIterationDemo
{
    // This example creates a workbook, adds a rectangle, oval, and textbox, then iterates the worksheet's Shapes collection. For each shape it extracts the Name, MsoDrawingType, Left, Top, Width, and Height properties and writes the information to the console before optionally saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample shapes to demonstrate the iteration
            // Rectangle shape
            Shape rect = sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);
            rect.Name = "MyRectangle";

            // Oval shape
            Shape oval = sheet.Shapes.AddOval(5, 0, 5, 0, 80, 120);
            oval.Name = "MyOval";

            // TextBox shape
            Shape txtBox = sheet.Shapes.AddShape(MsoDrawingType.TextBox, 8, 0, 8, 0, 150, 60);
            txtBox.Name = "MyTextBox";

            // Iterate all shapes on the worksheet
            for (int i = 0; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];

                // Retrieve shape properties
                string name = shape.Name;
                MsoDrawingType type = shape.MsoDrawingType;
                int left = shape.Left;     // X coordinate (pixels) from the left edge of the worksheet
                int top = shape.Top;       // Y coordinate (pixels) from the top edge of the worksheet
                int width = shape.Width;   // Width in pixels
                int height = shape.Height; // Height in pixels

                // Output the details
                Console.WriteLine($"Shape {i}: Name=\"{name}\", Type={type}, Left={left}, Top={top}, Width={width}, Height={height}");
            }

            // Save the workbook (optional, just to keep the file valid)
            workbook.Save("ShapesInfo.xlsx");
        }
    }
}
