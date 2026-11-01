// Title: Add a rectangle shape with a hover tooltip to an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add rectangle shape with tooltip to Excel | How to set alternative text for a shape in Excel using Aspose.Cells .NET | Create Excel shape that shows a tooltip on mouse hover with Aspose.Cells | Example of adding a shape with hover comment in a .xlsx file using Aspose.Cells | Set shape tooltip programmatically in C# with Aspose.Cells API
// Tags: add rectangle shape Aspose.Cells .NET | set shape alternative text Excel C# | shape tooltip Aspose.Cells example | Excel shape hover comment using Aspose | Aspose.Cells create shape with tooltip | C# Aspose.Cells shape properties

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a rectangle shape on the first worksheet, assigns visible text, sets the shape's AlternativeText property (displayed as a tooltip on mouse hover), and saves the workbook as ShapeWithTooltip.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: shape type, upper left row, upper left column, upper left offset (pixels), upper left offset (pixels), width (pixels), height (pixels)
        Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 80);

        // Optional: set visible text inside the shape
        shape.Text = "Hover over me";

        // Set the alternative text which Excel shows as a tooltip on mouse hover
        shape.AlternativeText = "This is a tooltip displayed when hovering over the shape.";

        // Save the workbook to a file
        workbook.Save("ShapeWithTooltip.xlsx", SaveFormat.Xlsx);
    }
}
