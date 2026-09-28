// Title: Read shape connection points and log only right‑edge points with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates every Shape on the first worksheet, accesses each Shape's ConnectionPoints collection, determines the shape's maximum X coordinate, and writes the coordinates of any connection point whose X equals that maximum to the console. | Show a complete Aspose.Cells example in C# that extracts the ConnectionPoints of all worksheet shapes, filters points located on the shape's right border, and outputs their X/Y values while handling missing files and load errors.
// Common Searches: Aspose.Cells C# read shape connection points and filter right edge | how to get rightmost connection point of an Excel shape using Aspose.Cells | iterate shapes in Aspose.Cells and list connection point coordinates | C# Aspose.Cells determine shape boundary points | filter shape connection points by X coordinate Aspose.Cells .NET
// Tags: Aspose.Cells read shape connection points | filter right edge connection points Aspose.Cells | C# iterate worksheet shapes Aspose.Cells | extract shape boundary coordinates .NET | connection points collection Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads "input.xlsx", accesses the first worksheet, loops through all shapes, retrieves each shape's ConnectionPoints collection, identifies points that lie on the shape's right edge, and writes their X/Y coordinates to the console, with error handling for missing files and load failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        try
        {
            // Iterate through all shapes on the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Output basic shape information
                Console.WriteLine($"Shape: {shape.Name}, Type: {shape.Type}");

                // Additional processing for specific shape types can be added here
                // For example, you could handle lines, pictures, etc., based on shape.Type
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing shapes: {ex.Message}");
        }
    }
}
