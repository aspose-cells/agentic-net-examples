// Title: How to ungroup Excel shapes while preserving their formatting using Aspose.Cells for .NET (C#)
// Common Searches: C# code to ungroup grouped shapes in an Excel workbook with Aspose.Cells and keep original style | Aspose.Cells .NET ungroup GroupShape objects without losing shape formatting | Remove shape groups from a worksheet using Aspose.Cells while retaining colors and borders | Example of iterating ShapeCollection to ungroup shapes in Aspose.Cells C# | How to save an Excel file after ungrouping shapes with Aspose.Cells
// Tags: Aspose.Cells ungroup GroupShape C# | preserve shape formatting Aspose.Cells | Excel worksheet shape collection manipulation Aspose.Cells | iterate shapes to ungroup Aspose.Cells | save workbook after ungrouping shapes Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsUngroupShapes
{
    // The example loads an Excel workbook, accesses the first worksheet, collects all GroupShape objects from the worksheet's Shapes collection, calls Ungroup() on each group to restore the individual shapes while automatically preserving their original formatting, and finally saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of shapes on the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Collect all group shapes first to avoid modifying the collection while iterating
                List<GroupShape> groupShapes = new List<GroupShape>();
                foreach (Shape shape in shapes)
                {
                    if (shape.IsGroup && shape is GroupShape groupShape)
                    {
                        groupShapes.Add(groupShape);
                    }
                }

                // Ungroup each group shape; the ungrouped shapes retain their original formatting automatically
                foreach (GroupShape groupShape in groupShapes)
                {
                    try
                    {
                        // Ungroup adds the inner shapes back to the worksheet; no return value.
                        groupShape.Ungroup();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to ungroup a shape: {ex.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the ungrouped shapes
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
