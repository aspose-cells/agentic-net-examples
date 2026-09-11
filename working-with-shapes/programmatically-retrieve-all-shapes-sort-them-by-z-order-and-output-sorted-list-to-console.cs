// Title: How to retrieve all worksheet shapes, sort them by Z‑order, and display the sorted list using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, iterates through every worksheet, collects each shape's name, sheet name, and ZOrderPosition, sorts the collection by ZOrderPosition ascending, and prints the results to the console. | Modify the shape enumeration to include only picture shapes, sort them by ZOrderPosition descending, and output the sorted information to the console. | Extend the example to write the sorted shape details (sheet, shape name, Z-order) to a CSV file instead of the console.
// Common Searches: aspnet retrieve shape Z-order from Excel using Aspose.Cells | c# list all shapes in workbook and sort by Z-order Aspose.Cells | how to get shape name and Z-order position across worksheets with Aspose.Cells for .NET | sorting Excel shapes by Z-order programmatically in C# | display sorted shape information from multiple sheets using Aspose.Cells
// Tags: Aspose.Cells enumerate worksheet shapes | C# sort shapes by ZOrderPosition | Aspose.Cells retrieve shape Z-order | console output sorted Excel shapes | export shape list to CSV Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook with Aspose.Cells, gathers each shape's sheet name, shape name, and ZOrderPosition from all worksheets, sorts the entries by Z-order ascending, and writes the sorted list to the console.
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

            // Collect information about every shape in the workbook
            List<(string SheetName, string ShapeName, int ZOrder)> shapeInfos = new List<(string, string, int)>();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Store sheet name, shape name and its Z-order position
                    shapeInfos.Add((sheet.Name, shape.Name, shape.ZOrderPosition));
                }
            }

            // Sort the shapes by their Z-order (ascending)
            List<(string SheetName, string ShapeName, int ZOrder)> sortedShapes = shapeInfos
                .OrderBy(info => info.ZOrder)
                .ToList();

            // Output the sorted list to the console
            Console.WriteLine("Shapes sorted by Z-order:");
            foreach (var info in sortedShapes)
            {
                Console.WriteLine($"Sheet: {info.SheetName}, Shape: {info.ShapeName}, Z-Order: {info.ZOrder}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
