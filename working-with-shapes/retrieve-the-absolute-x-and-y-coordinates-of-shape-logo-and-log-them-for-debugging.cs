// Title: Get absolute X/Y pixel coordinates of a named shape (Logo) with Aspose.Cells for .NET
// Description: Creates a workbook, adds a rectangle named "Logo", locates the shape via its Name property, calls GetActualBox() to obtain the left (X) and top (Y) pixel positions, writes the values to the console for debugging, and saves the workbook.
// Keywords: Aspose.Cells GetActualBox | shape position .NET | retrieve shape coordinates | absolute pixel location | named shape Logo | worksheet shape coordinates | Aspose.Cells debugging | C# shape bounding box
// Common Searches: Aspose.Cells get shape X coordinate | How to read shape Y position in .NET | GetActualBox example C# | Find shape by name Aspose.Cells | Log shape pixel location
// Developer Intent: Extract and output the absolute pixel X and Y positions of the worksheet shape called "Logo".
// Use Cases: Diagnose layout problems by printing the exact pixel location of a shape. | Align or distribute multiple shapes programmatically based on their absolute positions. | Export shape coordinates to another system for reporting or further processing.
// AI Prompts: Show me C# code that retrieves the absolute X and Y pixel coordinates of a shape named "Logo" using Aspose.Cells and logs them. | Provide an example of iterating through worksheet shapes, finding a specific one, and using GetActualBox to get its bounding box. | Explain how to use the X/Y values returned by GetActualBox to align other objects on an Excel sheet with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a rectangle named "Logo", locates the shape via its Name property, calls GetActualBox() to obtain the left (X) and top (Y) pixel positions, writes the values to the console for debugging, and saves the workbook.
    public class RetrieveShapeCoordinates
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape and name it "Logo"
                // Parameters: topRow, top, leftColumn, left, height, width
                Shape logoShape = worksheet.Shapes.AddRectangle(5, 5, 10, 10, 100, 50);
                logoShape.Name = "Logo";

                // Retrieve the shape named "Logo"
                Shape targetShape = null;
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.Name == "Logo")
                    {
                        targetShape = shape;
                        break;
                    }
                }

                if (targetShape != null)
                {
                    // Get the actual bounding box (left, top, right, bottom) in pixels
                    float[] box = targetShape.GetActualBox();

                    // Log absolute X (left) and Y (top) coordinates
                    Console.WriteLine($"Logo absolute X (pixels): {box[0]}");
                    Console.WriteLine($"Logo absolute Y (pixels): {box[1]}");
                }
                else
                {
                    Console.WriteLine("Shape named 'Logo' was not found.");
                }

                // Save the workbook (optional, just to complete the lifecycle)
                string outputPath = "RetrieveShapeCoordinates.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveShapeCoordinates.Run();
        }
    }
}
