// Title: Group and reposition rectangle shapes in an Excel file using Aspose.Cells for .NET (C#)
// Description: This C# sample creates a workbook, adds three rectangle shapes to the first worksheet, merges them into a GroupShape via ShapeCollection.Group, adjusts the group's Left and Top offsets to shift the entire collection, and saves the file as GroupedRectanglesDemo.xlsx.
// Keywords: Aspose.Cells C# group shapes | ShapeCollection.Group method | GroupShape positioning | move multiple shapes together | rectangle shape grouping Aspose | Excel shape container .NET
// Common Searches: Aspose.Cells how to group shapes | C# move grouped shapes in Excel | ShapeCollection.Group usage example | Set position of GroupShape Aspose.Cells | Combine rectangle shapes into one object
// Developer Intent: Combine several rectangle objects into a single GroupShape and modify its coordinates as a unified element.
// Use Cases: Create composite diagram components (e.g., flowchart blocks) that can be placed as a unit. | Align and shift related shapes together when generating automated reports. | Apply a single transformation—such as move or resize—to multiple shapes without handling each individually.
// AI Prompts: Generate code to ungroup a GroupShape and access its child shapes in Aspose.Cells. | Show how to add a caption textbox to an existing GroupShape using C#. | Demonstrate resizing a GroupShape while preserving the relative layout of its inner shapes in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGroupingDemo
{
    // This C# sample creates a workbook, adds three rectangle shapes to the first worksheet, merges them into a GroupShape via ShapeCollection.Group, adjusts the group's Left and Top offsets to shift the entire collection, and saves the file as GroupedRectanglesDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the shapes collection from the worksheet
            ShapeCollection shapes = worksheet.Shapes;

            // Add first rectangle shape
            // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
            Shape rect1 = shapes.AddRectangle(2, 0, 2, 0, 50, 80);
            rect1.Name = "Rectangle1";

            // Add second rectangle shape
            Shape rect2 = shapes.AddRectangle(6, 0, 2, 0, 50, 80);
            rect2.Name = "Rectangle2";

            // Add third rectangle shape (optional, demonstrates grouping more than two)
            Shape rect3 = shapes.AddRectangle(10, 0, 2, 0, 50, 80);
            rect3.Name = "Rectangle3";

            // Prepare an array of shapes to be grouped
            Shape[] shapesToGroup = new Shape[] { rect1, rect2, rect3 };

            // Group the shapes using ShapeCollection.Group (feature rule)
            GroupShape group = shapes.Group(shapesToGroup);
            group.Name = "MyRectangleGroup";

            // Move the entire group by setting its position
            // For example, shift 100 pixels to the right and 50 pixels down
            group.Left += 100;
            group.Top += 50;

            // Save the workbook (save rule)
            workbook.Save("GroupedRectanglesDemo.xlsx");
        }
    }
}
