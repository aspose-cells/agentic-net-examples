// Title: How to configure each shape to move and resize with its cell in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, iterates over all shapes on the first worksheet, and sets each shape's Placement property to MoveAndSize so they follow cell changes. | Create a reusable C# method that accepts a Worksheet object and applies MoveAndSize placement to every shape, then returns the modified worksheet. | Write C# that, after linking shapes to cells with MoveAndSize, updates each shape's displayed text to match the value of the cell it is anchored to before saving the workbook.
// Common Searches: Aspose.Cells C# set shape placement to MoveAndSize for dynamic resizing | link Excel shapes to cells so they move with cell changes using Aspose.Cells .NET | how to make shapes follow cell size in Aspose.Cells workbook | C# iterate worksheet shapes and set placement type MoveAndSize Aspose.Cells
// Tags: Aspose.Cells shape placement MoveAndSize | C# iterate worksheet shapes Aspose.Cells | dynamic shape repositioning Excel .NET | link Excel shape to cell movement Aspose | shape move and size with cell Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing workbook, loops through every shape on the first worksheet, sets each shape's Placement property to MoveAndSize so the shape moves and resizes together with its underlying cell, and then saves the updated file.
class ShapeLinker
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            for (int i = 0; i < sheet.Shapes.Count; i++)
            {
                Shape shape = sheet.Shapes[i];

                // Set the shape to move and size with the cell (relative positioning)
                shape.Placement = PlacementType.MoveAndSize;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
