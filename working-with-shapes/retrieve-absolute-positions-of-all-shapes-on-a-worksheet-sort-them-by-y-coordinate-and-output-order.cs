// Title: C# example: Retrieve absolute positions of all worksheet shapes and sort them by vertical (Top) coordinate using Aspose.Cells
// Common Searches: Aspose.Cells C# get shape top and left coordinates from a worksheet | How to list all shapes on an Excel sheet and sort them by Y position with Aspose.Cells | C# code to retrieve shape collection from a worksheet and order by vertical placement | Sorting worksheet shapes by top coordinate using Aspose.Cells for .NET | Get shape collection and sort by Top property in Aspose.Cells C#
// Tags: Aspose.Cells shape top coordinate extraction | C# sort Excel shapes by vertical position | retrieve worksheet shape collection Aspose.Cells | order shapes by Y coordinate .NET | Aspose.Cells shape position sorting example

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, accesses the first worksheet, gathers all shapes, records each shape's Top (Y) and Left coordinates, sorts the shapes in ascending order of the Top value, prints the sorted order with shape names and positions, and saves the workbook unchanged.
class ShapePositionSorter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve all shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Store each shape with its Y (Top) coordinate
            List<(Shape shape, double top)> shapeList = new List<(Shape, double)>();
            foreach (Shape shp in shapes)
            {
                shapeList.Add((shp, shp.Top));
            }

            // Sort shapes by Y coordinate (ascending)
            List<(Shape shape, double top)> sortedShapes = shapeList
                .OrderBy(item => item.top)
                .ToList();

            // Output the sorted order
            int order = 1;
            foreach (var item in sortedShapes)
            {
                Console.WriteLine($"Order {order}: Shape Name = {item.shape.Name}, Top = {item.top}, Left = {item.shape.Left}");
                order++;
            }

            // Save the workbook (no modifications made, just preserving the file)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
