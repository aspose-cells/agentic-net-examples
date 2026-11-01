// Title: Convert all SmartArt shapes to GroupShape objects in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a workbook, iterates through every worksheet, finds SmartArt shapes, and replaces each with the GroupShape returned by Shape.GetResultOfSmartArt while keeping its original anchor. | Show how to remove a SmartArt shape after conversion and add the resulting GroupShape to the same cell range using Aspose.Cells. | Provide robust error‑handling patterns for GetResultOfSmartArt failures when processing multiple worksheets in a .NET application.
// Common Searches: Aspose.Cells C# convert SmartArt to GroupShape preserving cell position | How to replace SmartArt objects with groups in an Excel file using Aspose.Cells | GetResultOfSmartArt example for batch processing worksheets in .NET | C# code to iterate worksheets and convert SmartArt shapes in Aspose.Cells
// Tags: Aspose.Cells SmartArt GetResultOfSmartArt conversion | C# replace SmartArt with GroupShape in Excel | preserve shape anchor Aspose.Cells | batch SmartArt processing Aspose.Cells | error handling GetResultOfSmartArt C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an input Excel workbook, scans each worksheet for SmartArt shapes, converts each SmartArt to a GroupShape using Shape.GetResultOfSmartArt while preserving its original row/column anchor, removes the original SmartArt, inserts the new GroupShape at the same location, and saves the updated workbook.
class SmartArtConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Collect SmartArt shapes first to avoid modifying the collection during iteration
                List<Shape> smartArtShapes = new List<Shape>();
                foreach (Shape shape in sheet.Shapes)
                {
                    // Identify SmartArt shapes using the IsSmartArt property
                    if (shape.IsSmartArt)
                    {
                        smartArtShapes.Add(shape);
                    }
                }

                // Convert each SmartArt shape to a GroupShape
                foreach (Shape smartShape in smartArtShapes)
                {
                    try
                    {
                        // Get the resulting GroupShape from the SmartArt
                        GroupShape group = smartShape.GetResultOfSmartArt();

                        // Preserve original position
                        int topRow = group.UpperLeftRow;
                        int leftColumn = group.UpperLeftColumn;
                        int bottomRow = topRow;      // same row; size will be adjusted automatically
                        int rightColumn = leftColumn; // same column

                        // Remove the original SmartArt shape
                        sheet.Shapes.Remove(smartShape);

                        // Add the resulting GroupShape to the worksheet using the original anchor positions
                        sheet.Shapes.AddCopy(group, topRow, leftColumn, bottomRow, rightColumn);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to convert SmartArt shape on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
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
