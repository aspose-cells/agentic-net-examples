// Title: Export each shape from an Excel worksheet to individual SVG files using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through Worksheet.Shapes, clones each shape into a fresh workbook, and saves it as an SVG file with Aspose.Cells. | Show how to keep a shape's original row and column indices while saving it as SVG using Aspose.Cells in C#.
// Common Searches: C# Aspose.Cells export Excel shape to SVG file | How to save each shape in an Excel sheet as separate SVG using Aspose.Cells | Copy Excel shape to new workbook and convert to SVG with Aspose.Cells .NET | Maintain exact shape alignment during Excel to SVG conversion using Aspose.Cells | Batch export worksheet shapes to SVG with Aspose.Cells example
// Tags: export worksheet shapes to SVG Aspose.Cells | shape-to-SVG conversion .NET | copy shape to temporary workbook C# | preserve shape coordinates Aspose.Cells | batch SVG export of Excel shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeExportExample
{
    // The example loads a workbook, loops through all shapes on the first worksheet, clones each shape into a temporary workbook while retaining its original row/column position, and saves each temporary workbook as a separate SVG file (Shape_1.svg, Shape_2.svg, …) for easy reuse on web pages.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string sourcePath = "SourceWorkbook.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // first worksheet

                // Iterate through all shapes on the worksheet
                for (int i = 0; i < sourceSheet.Shapes.Count; i++)
                {
                    Shape shape = sourceSheet.Shapes[i];

                    // Create a temporary workbook to hold the single shape
                    Workbook tempWorkbook = new Workbook();
                    Worksheet tempSheet = tempWorkbook.Worksheets[0];

                    // Copy the shape into the temporary worksheet, preserving its position
                    tempSheet.Shapes.AddCopy(
                        shape,
                        shape.UpperLeftRow,
                        shape.UpperLeftColumn,
                        shape.UpperLeftRow,
                        shape.UpperLeftColumn);

                    // Define the output SVG file name
                    string svgFileName = $"Shape_{i + 1}.svg";

                    // Save the temporary workbook as SVG (contains only the copied shape)
                    tempWorkbook.Save(svgFileName, SaveFormat.Svg);
                }

                Console.WriteLine("All shapes have been exported to SVG files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
