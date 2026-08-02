// Title: C# – Enumerate Worksheet Shapes and Get Name, Type, and Absolute Coordinates with Aspose.Cells
// Description: This example creates a workbook, adds a rectangle, an oval, and a textbox, then loops through the worksheet's Shapes collection. For each shape it reads the Name, Type, UpperLeftRow, UpperLeftColumn, Top, Left, Height and Width properties, prints the details to the console, and saves the file as ShapesInfo.xlsx.
// Keywords: Aspose.Cells shape enumeration | C# iterate worksheet shapes | Excel shape coordinates .NET | retrieve shape name and type | shape absolute position Aspose.Cells | Aspose.Cells drawing objects | C# Excel shape properties
// Common Searches: how to list all shapes in an Aspose.Cells worksheet | get shape row column offsets with Aspose.Cells C# | enumerate drawing objects in Excel using Aspose.Cells | retrieve shape dimensions and type in .NET Excel file
// Developer Intent: List every shape on a worksheet and display its identifier, category, and exact location/size information.
// Use Cases: Create an audit of all graphics in a spreadsheet for compliance reporting. | Validate that charts, images, or text boxes are positioned in the correct cells before publishing. | Map shapes to data rows for dynamic visual updates driven by cell values.
// AI Prompts: Generate C# code with Aspose.Cells that exports each shape's name, type, row, column, top offset, left offset, height, and width to a CSV file. | Show how to filter the shape enumeration to process only rectangles and ovals while ignoring other shape types. | Explain how to assign custom names to shapes before iterating so the output contains meaningful identifiers.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeIterationDemo
{
    // This example creates a workbook, adds a rectangle, an oval, and a textbox, then loops through the worksheet's Shapes collection. For each shape it reads the Name, Type, UpperLeftRow, UpperLeftColumn, Top, Left, Height and Width properties, prints the details to the console, and saves the file as ShapesInfo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample shapes to demonstrate the iteration
            // Rectangle at row 2, column 2, height 100, width 200
            sheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);
            // Oval at row 5, column 5, height 80, width 120
            sheet.Shapes.AddOval(5, 0, 5, 0, 80, 120);
            // TextBox at row 8, column 3, height 60, width 150
            sheet.Shapes.AddTextBox(8, 0, 3, 0, 60, 150);

            // Iterate all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Shape name (may be empty if not set)
                string name = shape.Name;

                // Shape type (auto shape type)
                string type = shape.Type.ToString();

                // Absolute coordinates: upper-left cell row/column and pixel offsets
                int upperLeftRow = shape.UpperLeftRow;
                int upperLeftColumn = shape.UpperLeftColumn;
                int topOffset = shape.Top;      // pixel offset from the top of the upper-left cell
                int leftOffset = shape.Left;    // pixel offset from the left of the upper-left cell
                int height = shape.Height;      // height in pixels
                int width = shape.Width;        // width in pixels

                Console.WriteLine($"Shape Name: {name}");
                Console.WriteLine($"Shape Type: {type}");
                Console.WriteLine($"Location: Row {upperLeftRow}, Column {upperLeftColumn}");
                Console.WriteLine($"Offset: Top {topOffset}px, Left {leftOffset}px");
                Console.WriteLine($"Size: Height {height}px, Width {width}px");
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ShapesInfo.xlsx");
        }
    }
}
