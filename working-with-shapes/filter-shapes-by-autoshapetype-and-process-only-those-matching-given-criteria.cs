// Title: Filter and Modify AutoShapeType Shapes in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add various auto shapes, and process only those whose AutoShapeType matches a specified value (e.g., Rectangle). Matching shapes are recolored and given alternative text before the file is saved.
// Keywords: Aspose.Cells | C# | .NET | AutoShapeType | filter worksheet shapes | modify shape fill color | set shape alternative text | shape collection iteration | rectangle auto shape | shape processing example
// Common Searches: Aspose.Cells filter shapes by type | change fill color of specific auto shapes C# | select rectangle shapes in Aspose.Cells worksheet | iterate over worksheet shapes Aspose.Cells .NET | apply conditional styling to shapes Aspose.Cells
// Developer Intent: Identify shapes of a particular AutoShapeType and apply custom formatting or metadata.
// Use Cases: Highlight all rectangle shapes by changing their fill color. | Add descriptive alternative text to shapes of a chosen type for accessibility. | Apply type‑specific styling such as borders or shadows across a worksheet.
// AI Prompts: Write C# code using Aspose.Cells that loops through worksheet shapes and adds a red border to every Oval shape. | Create a method that finds all Diamond auto shapes in a workbook and sets a tooltip explaining their purpose. | Provide an example that filters shapes by AutoShapeType and exports their row/column positions to a CSV file.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, add various auto shapes, and process only those whose AutoShapeType matches a specified value (e.g., Rectangle). Matching shapes are recolored and given alternative text before the file is saved.
class ShapeFilterDemo
{
    static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add various auto shapes to the worksheet
        sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 2, 2, 100, 50, 0, 0);
        sheet.Shapes.AddAutoShape(AutoShapeType.Oval, 5, 2, 100, 50, 0, 0);
        sheet.Shapes.AddAutoShape(AutoShapeType.Diamond, 8, 2, 100, 50, 0, 0);

        // Define the shape type we want to process
        AutoShapeType targetType = AutoShapeType.Rectangle;

        // Iterate through all shapes and process only those matching the target type
        foreach (Shape shape in sheet.Shapes)
        {
            // Check the AutoShapeType of the shape
            if (shape.AutoShapeType == targetType)
            {
                // Example processing: change fill color and set alternative text
                shape.Fill.SolidFill.Color = Color.LightGreen;
                shape.AlternativeText = "Processed rectangle shape";
            }
        }

        // Save the workbook with the processed shapes
        string outputPath = "FilteredShapesDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
