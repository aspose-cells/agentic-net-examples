// Title: How to center text horizontally and vertically in a specific shape using Aspose.Cells for .NET (C#)
// AI Prompts: Set both TextHorizontalAlignment and TextVerticalAlignment of a worksheet shape to Center with Aspose.Cells in C#. | Create a rectangle shape when none exists and then align its text to the middle horizontally and vertically using Aspose.Cells for .NET. | Update an existing shape's text alignment to center on both axes and save the workbook with the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# align shape text to center horizontally and vertically | set text alignment for Excel shape using Aspose.Cells .NET | programmatically center text inside a shape in an Excel file with Aspose.Cells | C# Aspose.Cells example for TextHorizontalAlignment and TextVerticalAlignment | add rectangle shape if missing and center its text using Aspose.Cells
// Tags: Aspose.Cells TextHorizontalAlignment | Aspose.Cells TextVerticalAlignment | shape text alignment Aspose.Cells | add rectangle shape Aspose.Cells | modify shape properties .NET | Excel shape alignment C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads a workbook, ensures a shape exists (adding a rectangle if needed), sets the shape's TextHorizontalAlignment and TextVerticalAlignment properties to Center, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                Shape shape;

                // Ensure there is at least one shape; create one if none exist
                if (sheet.Shapes.Count > 0)
                {
                    shape = sheet.Shapes[0];
                }
                else
                {
                    // Add a simple rectangle shape as a placeholder
                    shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 30);
                }

                // Set text alignment to center both horizontally and vertically
                shape.TextHorizontalAlignment = TextAlignmentType.Center;
                shape.TextVerticalAlignment = TextAlignmentType.Center;

                // Save the modified workbook
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
