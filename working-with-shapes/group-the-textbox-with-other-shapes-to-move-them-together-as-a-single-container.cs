// Title: How to group a textbox with other shapes and move them as a single container using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# group textbox with rectangle and oval | move grouped shapes together in an Excel worksheet using Aspose.Cells | set freefloating placement for a GroupShape in Aspose.Cells | resize and reposition a GroupShape in an XLSX file with Aspose.Cells | create and save a shape group in a workbook using Aspose.Cells C#
// Tags: group shapes Aspose.Cells C# | freefloating placement groupshape Aspose.Cells | resize groupshape Excel workbook | move shape group worksheet Aspose.Cells | add textbox shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace GroupTextboxWithShapes
{
    // The example creates a new workbook, adds a rectangle, a textbox, and an oval to the first worksheet, groups them into a single GroupShape, configures its name, placement, caption, position, and size, and then saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the ShapeCollection from the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add a rectangle shape
            // Parameters: upperLeftRow, top (pixel offset), upperLeftColumn, left (pixel offset), height (pixels), width (pixels)
            Shape rectangle = shapes.AddRectangle(2, 0, 2, 0, 80, 150);
            rectangle.Fill.FillType = FillType.Solid;
            rectangle.Fill.SolidFill.Color = System.Drawing.Color.LightBlue;
            rectangle.Line.DashStyle = MsoLineDashStyle.Solid;
            rectangle.Line.Weight = 1;

            // Add a textbox shape
            Shape textBox = shapes.AddTextBox(6, 0, 2, 0, 60, 200);
            textBox.Text = "Grouped TextBox";
            textBox.Fill.FillType = FillType.Solid;
            textBox.Fill.SolidFill.Color = System.Drawing.Color.LightYellow;
            textBox.Line.DashStyle = MsoLineDashStyle.Solid;
            textBox.Line.Weight = 1;

            // Add an oval shape
            Shape oval = shapes.AddOval(12, 0, 2, 0, 70, 120);
            oval.Fill.FillType = FillType.Solid;
            oval.Fill.SolidFill.Color = System.Drawing.Color.LightGreen;
            oval.Line.DashStyle = MsoLineDashStyle.Solid;
            oval.Line.Weight = 1;

            // Group the three shapes together
            Shape[] shapesToGroup = new Shape[] { rectangle, textBox, oval };
            GroupShape group = shapes.Group(shapesToGroup);

            // Set some properties on the group shape
            group.Name = "MyShapeGroup";
            group.Placement = PlacementType.FreeFloating; // Allows moving as a single object
            group.Text = "Group Container"; // Optional caption for the group

            // Move the group to a new location (e.g., lower rows and columns)
            // The Top and Left properties are in pixels relative to the worksheet grid
            group.Top = 200;   // Move down
            group.Left = 300;  // Move right

            // Optionally resize the group
            group.Width = 400;
            group.Height = 250;

            // Save the workbook
            string outputPath = "GroupedShapesDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine("The rectangle, textbox, and oval have been grouped and moved together.");
        }
    }
}
