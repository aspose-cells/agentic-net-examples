// Title: Align multiple worksheet shapes to the left edge of a specific column with Aspose.Cells for .NET (C#)
// AI Prompts: Generate a C# method that accepts a Worksheet, a zero‑based column index, and a collection of shape names, then sets each shape's UpperLeftColumn property so its left edge aligns with the given column using Aspose.Cells. | Provide sample code that loads an Excel file, calls the alignment method to position shapes "Shape1" and "Shape2" at column B, and saves the modified workbook.
// Common Searches: aspnet cells align shape left edge to column programmatically | c# set UpperLeftColumn for multiple shapes in Aspose.Cells | how to move Excel shapes to a specific column using Aspose.Cells | batch align worksheet shapes to column B in .NET | shape alignment helper example Aspose.Cells C#
// Tags: shape left alignment UpperLeftColumn Aspose.Cells | batch shape positioning worksheet C# | align worksheet shapes to column Aspose.Cells | programmatic Excel shape movement .NET | shape alignment helper method C# | column‑based shape placement Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // Provides a C# helper that iterates over a list of shape names, sets each shape's UpperLeftColumn to align its left edge with a specified column, and demonstrates loading a workbook, aligning two shapes to column B, and saving the result.
    public static class ShapeAlignmentHelper
    {
        /// <param name="worksheet">The worksheet containing the shapes.</param>
        /// <param name="startColumn">Zero‑based index of the column to align to.</param>
        /// <param name="shapeNames">Names of the shapes to be aligned.</param>
        public static void AlignShapesToLeftEdge(Worksheet worksheet, int startColumn, IEnumerable<string> shapeNames)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (shapeNames == null) throw new ArgumentNullException(nameof(shapeNames));
            if (startColumn < 0) throw new ArgumentOutOfRangeException(nameof(startColumn));

            foreach (var name in shapeNames)
            {
                try
                {
                    Shape shape = worksheet.Shapes[name];
                    if (shape == null) continue;

                    // Align shape to the specified column.
                    shape.UpperLeftColumn = startColumn;

                    // Offsets are not required for basic left alignment; they are left at default values.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to align shape '{name}': {ex.Message}");
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook.
                Workbook wb = new Workbook(inputPath);
                Worksheet ws = wb.Worksheets[0];

                // Align the desired shapes to column B (index 1).
                ShapeAlignmentHelper.AlignShapesToLeftEdge(
                    ws,
                    startColumn: 1,
                    shapeNames: new[] { "Shape1", "Shape2" });

                // Save the modified workbook.
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
