// Title: Ungroup Shapes, Adjust Positions, and Save Workbook with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds two rectangle shapes, groups them, extracts the individual shapes, ungroups the group, shifts each shape 20 px right and 10 px down by updating the X and Y properties, and saves the workbook.
// Keywords: Aspose.Cells ungroup shape C# | modify shape position Aspose.Cells | move Excel shapes programmatically | GroupShape Ungroup Aspose.Cells .NET | adjust shape coordinates C# | Excel shape manipulation Aspose.Cells
// Common Searches: Aspose.Cells ungroup shape group C# | change position of individual shapes after ungrouping | C# code to shift Excel shapes using Aspose.Cells | how to move grouped rectangle shapes in .NET | save workbook after modifying shape locations Aspose.Cells
// Developer Intent: Break a GroupShape into its component shapes, reposition each shape, and persist the changes in the Excel file using Aspose.Cells for .NET.
// Use Cases: Apply distinct formatting to shapes that were originally grouped. | Re‑align shapes with data rows or columns after dynamic layout changes. | Prepare Excel reports where shape positions must be programmatically adjusted before distribution. | Automate diagram updates by ungrouping and repositioning shapes.
// AI Prompts: Write C# code with Aspose.Cells to ungroup a GroupShape, iterate its child shapes, apply a custom X/Y offset, and save the workbook. | Explain the steps to retrieve individual Shape objects from a GroupShape and modify their X and Y properties. | Show how to ensure shape position changes are saved without affecting other worksheet content in Aspose.Cells for .NET. | Provide a sample that groups two rectangles, ungroups them, moves each shape, and writes the file to disk.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds two rectangle shapes, groups them, extracts the individual shapes, ungroups the group, shifts each shape 20 px right and 10 px down by updating the X and Y properties, and saves the workbook.
    public class UngroupAndModifyShapes
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                ShapeCollection shapes = worksheet.Shapes;

                // Add two rectangle shapes to the worksheet
                Shape shape1 = shapes.AddRectangle(2, 0, 2, 0, 100, 50); // upperLeftRow, upperLeftColumn, upperLeftPixelRow, upperLeftPixelColumn, height, width
                Shape shape2 = shapes.AddRectangle(6, 0, 2, 0, 100, 50);

                // Group the two shapes
                GroupShape groupShape = shapes.Group(new Shape[] { shape1, shape2 });

                // Retrieve the individual shapes that are part of the group
                Shape[] groupedShapes = groupShape.GetGroupedShapes();

                // Ungroup the shapes
                groupShape.Ungroup();

                // After ungrouping, modify each shape's position
                // For demonstration, shift each shape 20 pixels to the right and 10 pixels down
                foreach (Shape s in groupedShapes)
                {
                    s.X += 20; // move right
                    s.Y += 10; // move down
                }

                // Save the workbook with the modified shapes
                string outputPath = "UngroupedAndModifiedShapes.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UngroupAndModifyShapes.Run();
        }
    }
}
