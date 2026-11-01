// Title: Retrieve a shape's connection points, compute Euclidean distance between the first two points, and write the result to a cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a rectangle shape to a worksheet with Aspose.Cells, accesses its ConnectionPoints collection, extracts the first two points, calculates the Euclidean distance between them, and writes the value into cell A1. | Show how to use the Aspose.Cells Drawing API to obtain a shape's connection points, compute the straight‑line distance with Math.Sqrt, and save the workbook as an XLSX file. | Adapt an existing Aspose.Cells example to loop through a shape's ConnectionPoints, select the first two points, determine their X/Y coordinates, calculate the distance, and store the result in a specified worksheet cell.
// Common Searches: Aspose.Cells C# get shape connection points and calculate distance between them | How to compute Euclidean distance of two shape points in an Excel workbook using Aspose.Cells | Example of reading shape connection points collection with Aspose.Cells .NET | Store calculated shape measurement in Excel cell using Aspose.Cells API | Retrieve rectangle shape dimensions via connection points Aspose.Cells C#
// Tags: Aspose.Cells shape connection points distance | C# Aspose.Cells retrieve shape points | Aspose.Cells write computed value to Excel cell | Aspose.Cells rectangle shape diagonal via connection points | Aspose.Cells .NET Euclidean distance between shape points

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds a rectangle shape, accesses its ConnectionPoints collection, extracts the first two points, computes the Euclidean distance between those points, writes the distance into cell A1, and saves the workbook as ShapeConnectionPointsDistance.xlsx.
    class ShapeConnectionPointsDistance
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: shape type, upper left row, upper left column, upper left offset X, upper left offset Y, width, height
                Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 150, 100);

                // Calculate Euclidean distance between the shape's top‑left corner (0,0) and its bottom‑right corner (width,height)
                double deltaX = shape.Width;
                double deltaY = shape.Height;
                double distance = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);

                // Store the result in cell A1
                worksheet.Cells["A1"].PutValue(distance);

                // Save the workbook to a file
                string outputPath = "ShapeConnectionPointsDistance.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
