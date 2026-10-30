// Title: Get absolute pixel and inch coordinates of a grouped shape in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to determine the absolute position of a GroupShape in points | convert shape position from points to inches with Aspose.Cells .NET | calculate pixel coordinates of a grouped shape in an Excel file using Aspose.Cells | recursive method to accumulate parent group offsets for shapes in Aspose.Cells | retrieve grouped shape location in inches and pixels in C# Aspose.Cells example
// Tags: group shape absolute coordinates Aspose.Cells | points to inches conversion Excel shapes .NET | pixel calculation from shape points Aspose.Cells | parent group offset aggregation C# | Excel worksheet shape positioning

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example loads an Excel workbook, finds the first GroupShape on the first worksheet, recursively adds its Left and Top values with those of any parent groups to compute absolute coordinates in points, converts the result to inches (1 inch = 72 points) and to pixels assuming 96 DPI, outputs the positions, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load workbook (use the provided load rule)
        // {LoadWorkbook: "input.xlsx"}
        Workbook workbook = new Workbook("input.xlsx");   // placeholder for the load rule

        Worksheet sheet = workbook.Worksheets[0];

        // Locate the first GroupShape on the worksheet
        GroupShape groupShape = null;
        foreach (Shape shape in sheet.Shapes)
        {
            if (shape is GroupShape)
            {
                groupShape = (GroupShape)shape;
                break;
            }
        }

        if (groupShape == null)
        {
            Console.WriteLine("No grouped shape found on the worksheet.");
            return;
        }

        // Calculate absolute position in points (1 point = 1/72 inch)
        double absoluteLeftPoints = 0;
        double absoluteTopPoints = 0;
        AccumulateAbsolutePosition(groupShape, ref absoluteLeftPoints, ref absoluteTopPoints);

        // Convert points to inches (1 inch = 72 points)
        double leftInches = absoluteLeftPoints / 72.0;
        double topInches  = absoluteTopPoints  / 72.0;

        // Convert points to pixels (assuming 96 DPI: 1 inch = 96 pixels)
        double leftPixels = absoluteLeftPoints * 96.0 / 72.0;
        double topPixels  = absoluteTopPoints  * 96.0 / 72.0;

        Console.WriteLine($"Absolute position (pixels): X = {leftPixels:F2} px, Y = {topPixels:F2} px");
        Console.WriteLine($"Absolute position (inches): X = {leftInches:F2}\", Y = {topInches:F2}\"");

        // Save workbook (use the provided save rule)
        // {SaveWorkbook: "output.xlsx"}
        workbook.Save("output.xlsx");   // placeholder for the save rule
    }

    // Recursively adds the position of the shape and all its parent groups
    static void AccumulateAbsolutePosition(Shape shape, ref double leftPoints, ref double topPoints)
    {
        // Shape.Left and Shape.Top are measured in points
        leftPoints += shape.Left;
        topPoints  += shape.Top;

        // If the shape belongs to a parent group, include the parent's offset
        if (shape.Group != null)
        {
            AccumulateAbsolutePosition(shape.Group, ref leftPoints, ref topPoints);
        }
    }
}
