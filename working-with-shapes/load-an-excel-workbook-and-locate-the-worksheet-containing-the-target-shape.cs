// Title: Find the worksheet that contains a specific shape in an Excel workbook with Aspose.Cells for .NET
// Description: Loads a workbook, scans each worksheet’s ShapeCollection for a shape named "MyRectangle", retrieves the shape’s Worksheet property, prints the worksheet name, and saves the file unchanged.
// Keywords: Aspose.Cells shape lookup | C# find shape worksheet | Excel shape parent sheet | search shape by name Aspose.Cells | iterate worksheets shapes .NET | Shape.Worksheet property | retrieve shape location
// Common Searches: how to get the worksheet of a shape using Aspose.Cells | find shape named MyRectangle across all sheets in a workbook | Aspose.Cells C# locate shape parent worksheet | search for a specific shape in Excel with Aspose.Cells
// Developer Intent: Identify the worksheet that holds a shape with a given name inside an Excel file.
// Use Cases: Verify the existence of a named shape before applying formatting or data binding. | Move or copy a shape after determining its current worksheet. | Log the locations of critical shapes for auditing or documentation.
// AI Prompts: Generate C# code with Aspose.Cells that finds a shape called 'Chart1' and returns its worksheet name. | Provide an example that iterates through all worksheets, locates a shape by name, and changes its fill color. | Write code that searches for a shape across a workbook and, if found, moves it to a sheet named 'Summary'.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads a workbook, scans each worksheet’s ShapeCollection for a shape named "MyRectangle", retrieves the shape’s Worksheet property, prints the worksheet name, and saves the file unchanged.
class LocateShapeWorksheet
{
    static void Main()
    {
        // Load an existing workbook from file (uses Workbook(string) constructor)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Name of the shape we want to locate
        string targetShapeName = "MyRectangle";

        Shape targetShape = null;
        Worksheet shapeWorksheet = null;

        // Iterate through all worksheets to find the shape by name
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the collection of shapes in the current worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Try to retrieve the shape using the name indexer
            Shape shape = shapes[targetShapeName];
            if (shape != null)
            {
                targetShape = shape;
                // Use Shape.Worksheet property to get the containing worksheet
                shapeWorksheet = shape.Worksheet;
                break;
            }
        }

        // Output the result
        if (targetShape != null && shapeWorksheet != null)
        {
            Console.WriteLine($"Shape '{targetShape.Name}' is located in worksheet '{shapeWorksheet.Name}'.");
        }
        else
        {
            Console.WriteLine($"Shape '{targetShapeName}' was not found in any worksheet.");
        }

        // Save the workbook (no modifications made) using the Save method
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
