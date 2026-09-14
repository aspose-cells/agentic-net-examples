// Title: Convert SmartArt shapes to GroupShape objects while preserving layout using Aspose.Cells for .NET
// AI Prompts: Replace every SmartArt shape in an Excel workbook with its GroupShape result from GetResultOfSmartArt, copying the original Left, Top, Width, and Height values. | Iterate through all worksheets, collect shapes to avoid modifying the collection, convert SmartArt to GroupShape, delete the source SmartArt, and save the updated workbook.
// Common Searches: Aspose.Cells GetResultOfSmartArt example for converting smartart to groupshape | how to keep shape size when converting smartart in C# Excel | replace smartart with group shape across all worksheets using Aspose.Cells | C# code to remove smartart after conversion in Excel file
// Tags: smartart to groupshape conversion aspose.cells | preserve shape dimensions during smartart conversion | iterate worksheet shapes collection safely | remove original smartart after conversion | save workbook after shape transformation aspose.cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates through each worksheet, collects all shapes, converts any SmartArt shape to a GroupShape using GetResultOfSmartArt, copies the original shape's position and size to the new GroupShape, removes the original SmartArt, and saves the modified workbook to the specified output path.
class SmartArtConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Collect all shapes first to avoid modifying the collection during iteration
                List<Shape> shapes = new List<Shape>();
                foreach (Shape shape in sheet.Shapes)
                {
                    shapes.Add(shape);
                }

                // Attempt to convert each shape that is a SmartArt
                foreach (Shape shape in shapes)
                {
                    try
                    {
                        // GetResultOfSmartArt returns a GroupShape for SmartArt shapes; throws otherwise
                        GroupShape group = shape.GetResultOfSmartArt();

                        // Preserve original position and size
                        group.Left = shape.Left;
                        group.Top = shape.Top;
                        group.Width = shape.Width;
                        group.Height = shape.Height;

                        // Remove the original SmartArt shape
                        sheet.Shapes.Remove(shape);
                    }
                    catch (Exception ex)
                    {
                        // Ignore shapes that are not SmartArt or any conversion errors
                        Console.WriteLine($"Failed to convert shape on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
