// Title: Ungroup grouped shapes in an Excel worksheet and retain original formatting with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, finds every GroupShape on the first worksheet, calls Ungroup() on each, and saves the workbook while preserving each shape's original formatting. | Show how to safely iterate over a worksheet's ShapeCollection, detect GroupShape instances, and ungroup them without modifying the collection during enumeration. | Provide a complete example that checks for the input file, processes the workbook, and outputs a new file where all grouped shapes are separated but keep their visual appearance.
// Common Searches: aspnet ungroup Excel shapes while keeping formatting | Aspose.Cells C# ungroup GroupShape preserve appearance | how to separate grouped shapes in a workbook using Aspose.Cells | C# code to ungroup shapes in an .xlsx file with Aspose.Cells | remove shape groups from worksheet without losing style Aspose
// Tags: ungroup GroupShape Aspose.Cells | preserve shape formatting Aspose.Cells | iterate ShapeCollection C# | process first worksheet Aspose.Cells | separate grouped shapes .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsUngroupExample
{
    // Loads an existing .xlsx workbook, iterates through the first worksheet's ShapeCollection, calls Ungroup() on any GroupShape to split it while retaining each child shape's original formatting, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "Input.xlsx";
                string outputPath = "Output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the collection of shapes on the worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Copy shapes to a list to avoid modifying the collection while iterating
                List<Shape> shapeList = new List<Shape>();
                foreach (Shape shp in shapes)
                {
                    shapeList.Add(shp);
                }

                // Iterate through the copied list and ungroup any GroupShape found
                foreach (Shape shp in shapeList)
                {
                    if (shp is GroupShape groupShape)
                    {
                        // Ungroup the shape; the original GroupShape is removed automatically
                        groupShape.Ungroup();
                    }
                }

                // Save the workbook with the shapes now ungrouped
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
