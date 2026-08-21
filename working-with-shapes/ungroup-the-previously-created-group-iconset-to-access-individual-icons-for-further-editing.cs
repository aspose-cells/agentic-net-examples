// Title: Ungroup a GroupShape and edit individual icons using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add rectangle shapes, group them, ungroup the GroupShape, modify each shape's fill color, and save the file with Aspose.Cells C# API.
// Keywords: Aspose.Cells | .NET | C# | GroupShape | Ungroup | shape editing | icon set | fill color | Excel workbook | worksheet shapes
// Common Searches: Aspose.Cells ungroup GroupShape C# | how to edit individual icons after grouping in Aspose.Cells | change shape fill color after Ungroup Aspose.Cells | C# code to ungroup shapes in Excel using Aspose.Cells
// Developer Intent: Separate a grouped shape to access and modify each component shape.
// Use Cases: Ungroup a grouped icon set to recolor each icon individually. | Retrieve and reposition shapes after ungrouping for custom layout. | Apply distinct formatting (fill, border, text) to shapes once they are ungrouped.
// AI Prompts: Write C# code with Aspose.Cells that ungroups a GroupShape and sets a unique border style for each shape. | Show how to loop through the shapes returned by Ungroup() and add a text label to any text box within the group.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add rectangle shapes, group them, ungroup the GroupShape, modify each shape's fill color, and save the file with Aspose.Cells C# API.
    public class UngroupIconSetDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add two rectangle shapes (simulating icons)
            Shape shape1 = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 50);
            Shape shape2 = worksheet.Shapes.AddRectangle(0, 0, 3, 0, 100, 50);

            // Group the two shapes into a GroupShape
            GroupShape groupShape = worksheet.Shapes.Group(new Shape[] { shape1, shape2 });

            // Ungroup the previously created group to access individual shapes
            groupShape.Ungroup();

            // After ungrouping, edit the individual shapes (e.g., change fill colors)
            shape1.FillFormat.ForeColor = Color.Yellow;
            shape2.FillFormat.ForeColor = Color.LightBlue;

            // Save the workbook
            string outputPath = "UngroupedIconSetDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
