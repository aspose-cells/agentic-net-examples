// Title: Group a textbox, rectangle, and ellipse into a single shape container and reposition it to a target cell with Aspose.Cells for .NET (C#)
// AI Prompts: Create a textbox, rectangle, and oval on a worksheet, combine them into a shape group, and set the group's UpperLeftRow and UpperLeftColumn to move it to cell L12. | Show C# code that uses Aspose.Cells to add multiple shapes, group them, and relocate the group to a specific row and column.
// Common Searches: Aspose.Cells C# how to group multiple shapes and move the group to a specific cell | example of grouping a textbox with rectangle and ellipse in Aspose.Cells | move shape group to row 12 column 12 using Aspose.Cells .NET | C# code to create shape container from different drawing types in Aspose.Cells
// Tags: Aspose.Cells shape grouping C# | relocate shape group to cell Aspose.Cells | textbox rectangle ellipse combined shape Aspose.Cells | shape container repositioning .NET | save workbook with grouped shapes Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates adding a textbox, rectangle, and ellipse to a worksheet, grouping them into a single shape, moving the group to row 12/column 12, and saving the workbook as GroupedShapes.xlsx using Aspose.Cells for .NET.
class GroupShapesExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox shape
            Shape textBox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 100, 200);
            textBox.Text = "Grouped TextBox";

            // Add a rectangle shape
            Shape rectangle = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 5, 0, 0, 100, 150);
            rectangle.Line.Weight = 1.0;
            rectangle.Line.DashStyle = MsoLineDashStyle.Solid;

            // Add an ellipse shape
            Shape ellipse = sheet.Shapes.AddShape(MsoDrawingType.Oval, 8, 8, 0, 0, 80, 120);
            ellipse.Line.Weight = 1.0;
            ellipse.Line.DashStyle = MsoLineDashStyle.Solid;

            // Group the shapes
            Shape[] shapesToGroup = new Shape[] { textBox, rectangle, ellipse };
            Shape group = sheet.Shapes.Group(shapesToGroup);

            // Move the grouped container to a new location (row 12, column 12)
            group.UpperLeftRow = 12;
            group.UpperLeftColumn = 12;

            // Save the workbook
            string outputPath = "GroupedShapes.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
