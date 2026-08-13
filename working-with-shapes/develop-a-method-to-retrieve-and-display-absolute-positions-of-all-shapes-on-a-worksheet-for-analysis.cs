// Title: C# – Get absolute pixel positions and dimensions of all worksheet shapes with Aspose.Cells
// Description: Loads an Excel workbook, accesses the first worksheet, iterates its ShapeCollection, and prints each shape’s name, type, X/Y pixel offsets, width and height. Includes file‑existence verification and robust exception handling.
// Keywords: Aspose.Cells shape coordinates | C# get shape position Excel | retrieve shape dimensions Aspose.Cells | list worksheet shapes .NET | shape X Y pixels Aspose.Cells | Excel shape analysis C#
// Common Searches: Aspose.Cells how to read shape X and Y coordinates | C# list all shapes in Excel worksheet | Get shape width and height using Aspose.Cells | Retrieve shape positions in pixels from Excel file | Aspose.Cells shape collection example
// Developer Intent: Extract and display the absolute pixel location and size of every shape on a worksheet.
// Use Cases: Create a layout audit report of all shapes for UI design validation. | Programmatically align, reposition, or resize shapes based on their current coordinates. | Export shape position and size data to CSV or JSON for automated testing of Excel templates.
// AI Prompts: Write a method that returns a collection of objects containing shape name, type, X, Y, width, and height using Aspose.Cells for .NET. | Extend the sample to also capture each shape's Z‑order and the address of its top‑left anchor cell. | Create a function that shifts every shape by a specified X/Y offset while preserving its original size.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeAnalysis
{
    // Loads an Excel workbook, accesses the first worksheet, iterates its ShapeCollection, and prints each shape’s name, type, X/Y pixel offsets, width and height. Includes file‑existence verification and robust exception handling.
    public class ShapePositionRetriever
    {
        /// <param name="filePath">Path to the Excel file to be analyzed.</param>
        public static void RetrieveAndDisplayShapePositions(string filePath)
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found -> {filePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Get the collection of shapes on the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Iterate through each shape and output its absolute position and size
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Absolute position in pixels from the worksheet's top‑left corner
                    int x = shape.X;      // Horizontal offset
                    int y = shape.Y;      // Vertical offset

                    // Size in pixels
                    int width = shape.Width;
                    int height = shape.Height;

                    // Output details
                    Console.WriteLine($"Shape {i + 1}:");
                    Console.WriteLine($"  Name   : {shape.Name}");
                    Console.WriteLine($"  Type   : {shape.Type}");
                    Console.WriteLine($"  X      : {x} px");
                    Console.WriteLine($"  Y      : {y} px");
                    Console.WriteLine($"  Width  : {width} px");
                    Console.WriteLine($"  Height : {height} px");
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            // Path to the workbook containing shapes
            string inputFile = "ShapesDemo.xlsx";

            try
            {
                RetrieveAndDisplayShapePositions(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
