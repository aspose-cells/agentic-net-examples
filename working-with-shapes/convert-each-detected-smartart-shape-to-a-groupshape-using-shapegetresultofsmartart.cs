// Title: Convert SmartArt shapes to GroupShape objects in an Excel workbook using Aspose.Cells Shape.GetResultOfSmartArt (C#)
// Common Searches: C# Aspose.Cells replace SmartArt with GroupShape in an existing workbook | How to extract SmartArt as a GroupShape using Shape.GetResultOfSmartArt in Aspose.Cells | Iterate all worksheets and convert every SmartArt shape to a GroupShape with Aspose.Cells .NET | Remove original SmartArt after converting it to a GroupShape in Aspose.Cells | Adjust size and position of a GroupShape after SmartArt conversion using Aspose.Cells
// Tags: SmartArt to GroupShape conversion Aspose.Cells | Shape.GetResultOfSmartArt C# example | Iterate worksheet shapes Aspose.Cells | Remove SmartArt after conversion Aspose.Cells | Excel workbook SmartArt handling Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, scans each worksheet for SmartArt shapes, converts each SmartArt to a GroupShape via Shape.GetResultOfSmartArt, optionally syncs its dimensions, removes the original SmartArt, and saves the updated file.
class SmartArtConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Collect SmartArt shapes first to avoid modifying the collection while iterating
                List<Shape> smartArtShapes = new List<Shape>();

                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.IsSmartArt)
                        smartArtShapes.Add(shape);
                }

                // Convert each SmartArt shape to a GroupShape
                foreach (Shape smartArt in smartArtShapes)
                {
                    // Convert SmartArt; the resulting GroupShape is automatically added to the worksheet
                    GroupShape resultGroup = smartArt.GetResultOfSmartArt();

                    // Optional: you can adjust the position/size of resultGroup here if needed
                    // resultGroup.Width = smartArt.Width;
                    // resultGroup.Height = smartArt.Height;
                    // resultGroup.Left = smartArt.Left;
                    // resultGroup.Top = smartArt.Top;

                    // Remove the original SmartArt shape from the worksheet
                    sheet.Shapes.Remove(smartArt);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
