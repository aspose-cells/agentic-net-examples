// Title: Adjust shape Z‑order to place one shape directly above another using Aspose.Cells for .NET
// Description: Creates a workbook, adds two rectangle shapes, sets explicit ZOrderPosition values, moves the second shape to be just above the first, and saves the file.
// Keywords: Aspose.Cells shape Z order | C# ZOrderPosition | move shape above another Aspose.Cells | Excel shape layering .NET | adjust shape Z index
// Common Searches: Aspose.Cells change shape ZOrderPosition | place one shape above another in Excel with Aspose.Cells | reorder overlapping shapes .NET | how to set shape layering in Aspose.Cells | bring shape to front programmatically Aspose
// Developer Intent: Set a shape's Z‑order so it appears directly above a specific existing shape.
// Use Cases: Control visibility when rectangles overlap in a report. | Keep a watermark behind data while a callout stays on top. | Position a label shape immediately above a chart for clear annotation.
// AI Prompts: Generate C# code that moves a given Aspose.Cells Shape to a ZOrderPosition just above another specified shape. | Show how to reorder multiple shapes in a worksheet so each shape is placed one level higher than the previous one. | Write a method that accepts two Shape objects and an offset, then adjusts the second shape's ZOrderPosition relative to the first using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Creates a workbook, adds two rectangle shapes, sets explicit ZOrderPosition values, moves the second shape to be just above the first, and saves the file.
    public class AdjustZOrder
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add two rectangle shapes
                Shape shapeA = worksheet.Shapes.AddRectangle(5, 5, 100, 100, 0, 0);   // Existing shape
                Shape shapeB = worksheet.Shapes.AddRectangle(50, 50, 100, 100, 0, 0); // Shape to move

                // Ensure initial Z-order positions (optional, default is sequential)
                shapeA.ZOrderPosition = 0; // bottom
                shapeB.ZOrderPosition = 1; // top

                // Move shapeB to be just above shapeA
                shapeB.ZOrderPosition = shapeA.ZOrderPosition + 1;

                // Define output file path
                string outputPath = "ZOrderAdjusted.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
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
            AdjustZOrder.Run();
        }
    }
}
