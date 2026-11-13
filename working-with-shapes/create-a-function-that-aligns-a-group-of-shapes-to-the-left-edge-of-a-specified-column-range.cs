// Title: C# method to align a group of worksheet shapes to the left edge of a specific column using Aspose.Cells
// Common Searches: how to move several shapes to the same column in Aspose.Cells C# | Aspose.Cells set shape UpperLeftColumn without changing size | align shapes to column index programmatically in .NET workbook | batch update shape positions in Excel worksheet using Aspose.Cells | C# code to align shapes to left edge of column range in Aspose.Cells
// Tags: shape UpperLeftColumn alignment Aspose.Cells | batch shape positioning worksheet C# | preserve shape dimensions while moving Aspose.Cells | missing shape handling worksheet Aspose.Cells | column-based shape alignment Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeAlignmentDemo
{
    // The AlignShapesToLeftColumn method receives a Workbook, worksheet name, start column index, and a list of shape names, then repositions each existing shape so its UpperLeftColumn matches the specified column while retaining its original lower‑right corner, handling null arguments, missing worksheets, and absent shapes without interrupting the process.
    public class ShapeAlignmentHelper
    {
        // Aligns the left edge of the specified shapes to the left edge of the start column in the given worksheet.
        public static void AlignShapesToLeftColumn(Workbook workbook, string sheetName, int startColumn, List<string> shapeNames)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrEmpty(sheetName)) throw new ArgumentException("Sheet name cannot be null or empty.", nameof(sheetName));
            if (shapeNames == null) throw new ArgumentNullException(nameof(shapeNames));

            try
            {
                // Get the worksheet safely; indexer returns null if not found.
                Worksheet sheet = workbook.Worksheets[sheetName];
                if (sheet == null)
                {
                    Console.Error.WriteLine($"Worksheet \"{sheetName}\" not found.");
                    return;
                }

                foreach (string shapeName in shapeNames)
                {
                    if (string.IsNullOrEmpty(shapeName))
                        continue;

                    Shape shape = sheet.Shapes[shapeName];
                    if (shape == null)
                    {
                        // Shape not found; continue with next.
                        continue;
                    }

                    // Preserve the shape's lower‑right corner to keep its size unchanged.
                    int lowerRightRow = shape.LowerRightRow;
                    int lowerRightColumn = shape.LowerRightColumn;

                    // Move the shape so its upper‑left column aligns with startColumn.
                    shape.UpperLeftColumn = startColumn;
                    shape.LowerRightRow = lowerRightRow;
                    shape.LowerRightColumn = lowerRightColumn;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error aligning shapes: {ex.Message}");
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
                Workbook workbook;

                // Load existing workbook if present; otherwise create a new one.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // New workbook with a default worksheet.
                }

                // Example: align shapes named "Shape 1" and "Shape 2" on "Sheet1" to column index 2 (C).
                var shapeNames = new List<string> { "Shape 1", "Shape 2" };
                ShapeAlignmentHelper.AlignShapesToLeftColumn(workbook, "Sheet1", 2, shapeNames);

                // Save the modified workbook.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
