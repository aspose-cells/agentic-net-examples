using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtInfo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Cast to SmartArtShape to access SmartArt‑specific members if needed
                        SmartArtShape smartArt = shape as SmartArtShape;

                        // Log detailed information
                        Console.WriteLine("SmartArt Shape Details:");
                        Console.WriteLine($"  Name          : {shape.Name}");
                        Console.WriteLine($"  Type          : {shape.Type}");
                        Console.WriteLine($"  Position (X,Y): ({shape.Left}, {shape.Top})");
                        Console.WriteLine($"  Size (W,H)    : ({shape.Width}, {shape.Height})");
                        Console.WriteLine($"  Left (pixels) : {shape.Left}");
                        Console.WriteLine($"  Top (pixels)  : {shape.Top}");
                        Console.WriteLine($"  Width (pixels): {shape.Width}");
                        Console.WriteLine($"  Height (pixels): {shape.Height}");
                        Console.WriteLine();
                    }
                }
            }

            // Save the workbook (if any modifications were made)
            workbook.Save("output.xlsx");
        }
    }
}