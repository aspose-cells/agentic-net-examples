// Title: Group a TextBox with Rectangle and Oval using Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, adds a TextBox, a rectangle, and an oval to the first worksheet, groups them into a single GroupShape named "MyGroupedShapes", and saves the file as GroupedShapes.xlsx.
// Keywords: Aspose.Cells | C# | .NET | GroupShape | group shapes | textbox grouping | Excel shape grouping | Aspose.Cells example | shape container
// Common Searches: Aspose.Cells group textbox with other shapes | C# group rectangle and oval in Excel using Aspose | How to create a GroupShape in Aspose.Cells | Move multiple shapes together in an Excel workbook C# | Aspose.Cells GroupShape method usage
// Developer Intent: Combine a TextBox, rectangle, and oval into a single GroupShape so they can be moved and formatted as one unit.
// Use Cases: Design a diagram where annotation shapes stay aligned when repositioned. | Create a composite label for a chart that includes text and geometric shapes. | Apply uniform scaling or rotation to several related shapes with a single command.
// AI Prompts: Show C# code to ungroup a GroupShape created with Aspose.Cells. | Provide an example of adding a caption to an existing GroupShape in Aspose.Cells for .NET. | Explain how to iterate through individual shapes inside a GroupShape using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGroupTextboxExample
{
    // C# example that creates a workbook, adds a TextBox, a rectangle, and an oval to the first worksheet, groups them into a single GroupShape named "MyGroupedShapes", and saves the file as GroupedShapes.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Get the shapes collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Add a TextBox shape
            TextBox txtBox = shapes.AddTextBox(2, 0, 2, 0, 120, 40);
            txtBox.Text = "Sample TextBox";

            // Add a rectangle shape
            Shape rect = shapes.AddRectangle(5, 0, 5, 0, 100, 60);
            rect.Text = "Rectangle";

            // Add an oval shape
            Shape oval = shapes.AddOval(9, 0, 9, 0, 80, 80);
            oval.Text = "Oval";

            // Group the TextBox with the rectangle and oval
            Shape[] itemsToGroup = new Shape[] { txtBox, rect, oval };
            GroupShape group = shapes.Group(itemsToGroup);
            group.Name = "MyGroupedShapes";

            // Save the workbook
            workbook.Save("GroupedShapes.xlsx");
        }
    }
}
