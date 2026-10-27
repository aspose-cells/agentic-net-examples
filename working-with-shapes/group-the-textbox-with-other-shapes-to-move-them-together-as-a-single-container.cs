// Title: Moving a TextBox and a Rectangle together in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# move multiple shapes with same offset | how to shift a textbox and rectangle together in an Excel file using Aspose.Cells | C# programmatically adjust position of several worksheet shapes Aspose.Cells | move grouped shapes in Excel workbook with Aspose.Cells API | apply uniform top and left offsets to multiple shapes in Aspose.Cells
// Tags: move multiple worksheet shapes Aspose.Cells | adjust shape top left Aspose.Cells C# | textbox and rectangle positioning Aspose.Cells | batch shape offset Excel Aspose.Cells | shape coordinate manipulation Aspose.Cells | worksheet shape layout Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a TextBox and a Rectangle shape to the first worksheet, increments their Top and Left properties by the same values to move them together, and saves the file as GroupedShapesOutput.xlsx.
class GroupShapesExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a textbox to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            TextBox textbox = sheet.Shapes.AddTextBox(2, 2, 50, 50, 120, 40);
            textbox.Text = "Sample TextBox";

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, top offset, left offset, width, height
            Shape rectangle = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                4,    // upper left row
                4,    // upper left column
                80,   // top offset (pixels)
                80,   // left offset (pixels)
                150,  // width (pixels)
                80    // height (pixels)
            );

            // Move both shapes (30 pixels down, 50 pixels right)
            textbox.Top += 30;
            textbox.Left += 50;
            rectangle.Top += 30;
            rectangle.Left += 50;

            // Save the workbook with the shapes
            workbook.Save("GroupedShapesOutput.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
