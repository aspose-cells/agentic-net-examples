// Title: How to retrieve and filter shape connection points in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, obtains the first shape's connection points, and prints only the points where the X coordinate exceeds 100. | Create a method that safely invokes Shape.GetConnectionPoints, iterates the returned points, and returns a collection of points with X > 100 while handling missing X/Y properties. | Add robust error handling to a C# Aspose.Cells example that checks file existence, shape presence, and unsupported GetConnectionPoints, and logs a friendly message when no qualifying points are found.
// Common Searches: Aspose.Cells C# get connection points of a shape and filter by X coordinate | filter Excel shape connection points where X > 100 using Aspose.Cells | C# example for retrieving shape connection points with Aspose.Cells and handling unsupported methods | how to check if GetConnectionPoints is supported in Aspose.Cells for .NET | log shape connection points from first worksheet shape in Aspose.Cells C#
// Tags: Aspose.Cells shape GetConnectionPoints C# | filter connection points by X coordinate .NET | handle unsupported GetConnectionPoints Aspose.Cells | log shape connection points Excel workbook C# | retrieve first shape connection points Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that loads an Excel workbook with Aspose.Cells, accesses the first shape, retrieves its connection points via GetConnectionPoints, filters points where the X coordinate is greater than 100, and writes those coordinates to the console, including comprehensive error handling for missing files, shapes, or unsupported methods.
class Program
{
    static void Main()
    {
        string inputPath = @"C:\Path\To\YourWorkbook.xlsx";

        try
        {
            // Verify that the workbook file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one shape
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            // Retrieve the first shape
            Shape shape = sheet.Shapes[0];

            // Attempt to get connection points (may not be supported in older versions)
            dynamic points = null;
            try
            {
                points = shape.GetConnectionPoints();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unable to retrieve connection points: {ex.Message}");
            }

            if (points != null)
            {
                // Iterate through the connection points and log those with X > 100
                foreach (dynamic point in points)
                {
                    try
                    {
                        double x = point.X;
                        double y = point.Y;
                        if (x > 100)
                        {
                            Console.WriteLine($"Connection Point - X: {x}, Y: {y}");
                        }
                    }
                    catch
                    {
                        // Skip points that do not expose X/Y properties
                    }
                }
            }
            else
            {
                Console.WriteLine("No connection points available for this shape.");
            }

            // (Optional) Save the workbook if any changes were made
            // workbook.Save(@"C:\Path\To\ModifiedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
