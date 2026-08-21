// Title: Group and Move Multiple Rectangle Shapes with Aspose.Cells for .NET
// Description: Demonstrates how to add three rectangle shapes to a worksheet, combine them into a GroupShape using ShapeCollection.Group, rename the group, reposition it by adjusting Left and Top properties, and save the workbook as GroupedRectangles.xlsx.
// Keywords: Aspose.Cells shape grouping | C# GroupShape example | Excel rectangle group Aspose.Cells | ShapeCollection.Group method | move grouped shapes .NET | Aspose.Cells rectangle shapes
// Common Searches: group multiple shapes Aspose.Cells C# | move a group of shapes together in Excel using Aspose.Cells | how to create a GroupShape in Aspose.Cells .NET | Aspose.Cells rectangle shape grouping tutorial | shift grouped shapes left top Aspose.Cells
// Developer Intent: Create a GroupShape from several rectangle shapes and reposition the entire group with a single operation.
// Use Cases: Bundle related diagram elements for a single‑click layout adjustment after report generation. | Build a composite legend by grouping shapes and placing it in a fixed worksheet area. | Re‑align annotation rectangles in bulk when data rows are inserted or deleted.
// AI Prompts: Show how to ungroup a GroupShape and retrieve the original rectangle objects in Aspose.Cells. | Provide sample code that groups different shape types (rectangle, oval, line) and applies a common rotation. | Explain how to keep fill color and line style of grouped shapes when exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeGroupingDemo
{
    // Demonstrates how to add three rectangle shapes to a worksheet, combine them into a GroupShape using ShapeCollection.Group, rename the group, reposition it by adjusting Left and Top properties, and save the workbook as GroupedRectangles.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the shape collection of the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Add first rectangle shape
                // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
                Shape rect1 = shapes.AddRectangle(2, 0, 2, 0, 50, 80);
                rect1.Name = "Rect1";

                // Add second rectangle shape
                Shape rect2 = shapes.AddRectangle(6, 0, 2, 0, 50, 80);
                rect2.Name = "Rect2";

                // Add third rectangle shape
                Shape rect3 = shapes.AddRectangle(10, 0, 2, 0, 50, 80);
                rect3.Name = "Rect3";

                // Create an array of the shapes to be grouped
                Shape[] shapesToGroup = new Shape[] { rect1, rect2, rect3 };

                // Group the shapes using ShapeCollection.Group method
                GroupShape group = shapes.Group(shapesToGroup);
                group.Name = "MyRectangleGroup";

                // Move the whole group by setting its position
                // For example, shift 100 pixels to the right and 50 pixels down
                group.Left += 100;
                group.Top += 50;

                // Save the workbook
                workbook.Save("GroupedRectangles.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
