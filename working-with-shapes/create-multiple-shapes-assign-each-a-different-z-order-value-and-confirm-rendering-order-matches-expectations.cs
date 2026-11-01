// Title: Add three rectangle shapes with different Z‑order positions and confirm stacking order in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Create three rectangle shapes on a worksheet, set their ZOrderPosition to 1, 2, and 3 respectively, and print each shape’s name with its Z‑order value. | Retrieve all shapes from the worksheet, sort them by ZOrderPosition, and output the sequence from the bottom layer to the top layer. | Save the workbook after configuring the shape Z‑order and open the resulting .xlsx file to visually verify that the shape with ZOrderPosition 3 appears on top.
// Common Searches: Aspose.Cells C# assign shape stacking order in Excel | C# example sorting worksheet shapes by their drawing order using Aspose.Cells | how to check visual layering of shapes after saving workbook with Aspose.Cells | add multiple rectangle shapes with different layers in Aspose.Cells .NET
// Tags: Aspose.Cells shape ZOrderPosition assignment | insert rectangle shapes Aspose.Cells C# | order worksheet shapes by ZOrderPosition | verify Excel shape layering Aspose.Cells | save workbook with ordered shapes Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, adds three rectangle shapes with ZOrderPosition values 1‑3, prints each shape’s Z‑order, sorts the shapes to display the rendering sequence from bottom to top, and saves the file as ShapesZOrder.xlsx to confirm the visual stacking.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Add three rectangle shapes with different Z‑order
            // -------------------------------------------------

            // Shape 1 – will be at the bottom (Z‑order = 1)
            Shape shape1 = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row
                2,   // upper left column
                0,   // top offset (pixels)
                0,   // left offset (pixels)
                100, // height (pixels)
                200  // width (pixels)
            );
            shape1.Name = "Shape1";
            shape1.ZOrderPosition = 1;

            // Shape 2 – middle layer (Z‑order = 2)
            Shape shape2 = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                3,
                3,
                0,
                0,
                100,
                200
            );
            shape2.Name = "Shape2";
            shape2.ZOrderPosition = 2;

            // Shape 3 – topmost layer (Z‑order = 3)
            Shape shape3 = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                4,
                4,
                0,
                0,
                100,
                200
            );
            shape3.Name = "Shape3";
            shape3.ZOrderPosition = 3;

            // -------------------------------------------------
            // Verify the Z‑order values (rendering order)
            // -------------------------------------------------
            Console.WriteLine("Shapes and their Z‑order values:");
            foreach (Shape shp in sheet.Shapes)
            {
                Console.WriteLine($"{shp.Name} => ZOrder = {shp.ZOrderPosition}");
            }

            // Optional: confirm that higher Z‑order means on‑top
            // (by sorting ascending and printing the order)
            Console.WriteLine("\nRendering order from bottom to top:");
            List<Shape> ordered = new List<Shape>(sheet.Shapes);
            ordered.Sort((a, b) => a.ZOrderPosition.CompareTo(b.ZOrderPosition));
            foreach (Shape shp in ordered)
            {
                Console.WriteLine($"{shp.Name} (ZOrder={shp.ZOrderPosition})");
            }

            // -------------------------------------------------
            // Save the workbook to verify visual result
            // -------------------------------------------------
            string outputPath = "ShapesZOrder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to '{Path.GetFullPath(outputPath)}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
