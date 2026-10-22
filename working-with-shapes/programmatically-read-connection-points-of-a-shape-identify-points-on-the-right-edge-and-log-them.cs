// Title: How to read and log right‑edge connection points of Excel shapes using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, iterates every shape on the first worksheet, calls GetConnectionPoints(), and prints the points whose X coordinate is within 0.01 of the shape's Width. | Extend the shape loop to also capture left‑edge connection points by checking if the X coordinate is close to zero, and output their index and Y coordinate. | Add robust error handling: verify the input Excel file exists before loading, and wrap each shape's processing in a try‑catch that logs the shape name and exception without aborting the whole iteration.
// Common Searches: Aspose.Cells C# get shape connection points and filter by right edge | How to identify right‑most connection point of a shape in an Excel file using Aspose.Cells | Iterate over shapes in a worksheet and read connection point coordinates with Aspose.Cells for .NET | C# code to log connection points of Excel shapes that lie on the shape's right border | Aspose.Cells GetConnectionPoints example with tolerance for floating point
// Tags: Aspose.Cells shape connection point retrieval | filter right‑edge shape connection points C# | iterate worksheet shapes Aspose.Cells | compare shape width with connection point X C# | verify Excel file existence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook using Aspose.Cells, iterates all shapes on the first worksheet, obtains each shape's connection points via GetConnectionPoints(), and logs the points whose X coordinate matches the shape's width (right‑edge) within a 0.01 tolerance. It includes a file‑existence check and per‑shape exception handling to keep processing uninterrupted.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                try
                {
                    // Width of the shape in points
                    double shapeWidth = shape.Width;

                    Console.WriteLine($"Shape: {shape.Name}, Type: {shape.Type}");

                    // Retrieve connection points for the shape (type is ConnectionPointInfo[])
                    var connectionPoints = shape.GetConnectionPoints();

                    // Examine each connection point
                    foreach (var cpObj in connectionPoints)
                    {
                        // Use dynamic to access properties without explicit type reference
                        dynamic cp = cpObj;
                        double x = cp.X;
                        double y = cp.Y;
                        int index = cp.Index;

                        // Identify points that lie on the right edge (X coordinate ≈ shape width)
                        if (Math.Abs(x - shapeWidth) < 0.01) // tolerance for floating‑point precision
                        {
                            Console.WriteLine($"  Right‑edge connection point - Index: {index}, X: {x}, Y: {y}");
                        }
                    }
                }
                catch (Exception shapeEx)
                {
                    // Handle errors related to a specific shape without stopping the whole process
                    Console.WriteLine($"Error processing shape \"{shape.Name}\": {shapeEx.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
