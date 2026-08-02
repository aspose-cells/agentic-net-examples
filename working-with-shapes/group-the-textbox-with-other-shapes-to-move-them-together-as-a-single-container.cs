// Title: Group a TextBox with other shapes and move them as a single container using Aspose.Cells for .NET
// Description: Demonstrates how to add a TextBox, rectangle, and oval to a worksheet, combine them into a GroupShape, rename the group, and reposition the whole group by setting its Left and Top properties before saving the workbook.
// Keywords: Aspose.Cells group shapes | C# GroupShape example | move grouped shapes .NET | add TextBox Aspose.Cells | Excel shape container
// Common Searches: Aspose.Cells group TextBox with rectangle and oval | move grouped shapes together in Excel using C# | how to create a GroupShape in Aspose.Cells | set position of a grouped shape Aspose.Cells
// Developer Intent: Create a GroupShape that contains a TextBox, rectangle, and oval, then shift the entire group as one object.
// Use Cases: Bundle a label TextBox with decorative shapes to create a movable annotation on a dashboard. | Assemble a multi‑part logo from separate shapes and position it with a single coordinate change. | Align form fields composed of several shapes and adjust their layout by moving the group instead of each element.
// AI Prompts: Generate C# code with Aspose.Cells that groups a TextBox, rectangle, and oval and sets the group's Left and Top values. | Explain how to ungroup a GroupShape in Aspose.Cells while keeping the original shape positions intact. | Show how to add a hyperlink to a GroupShape that contains multiple shapes in an Excel file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a TextBox, rectangle, and oval to a worksheet, combine them into a GroupShape, rename the group, and reposition the whole group by setting its Left and Top properties before saving the workbook.
class GroupTextboxDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        ShapeCollection shapes = sheet.Shapes;

        // Add a TextBox shape
        TextBox txtBox = (TextBox)shapes.AddTextBox(2, 0, 2, 0, 150, 50);
        txtBox.Text = "Sample Text";
        txtBox.Name = "MyTextBox";

        // Add a rectangle shape
        Shape rect = shapes.AddRectangle(5, 0, 5, 0, 100, 80);
        rect.Name = "MyRectangle";

        // Add an oval shape
        Shape oval = shapes.AddOval(8, 0, 8, 0, 80, 80);
        oval.Name = "MyOval";

        // Group the TextBox with the rectangle and oval
        Shape[] groupItems = new Shape[] { txtBox, rect, oval };
        GroupShape group = shapes.Group(groupItems);
        group.Name = "MyGroup";

        // Move the grouped container as a single unit
        group.Left = 200; // horizontal offset in pixels
        group.Top = 100;  // vertical offset in pixels

        // Save the workbook
        workbook.Save("GroupedTextboxDemo.xlsx");
    }
}
