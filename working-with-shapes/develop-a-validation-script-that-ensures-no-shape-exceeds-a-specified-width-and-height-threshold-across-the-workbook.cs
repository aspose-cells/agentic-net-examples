// Title: C# script to validate and resize Excel shape dimensions using Aspose.Cells
// Description: Loads an Excel workbook, walks through every worksheet and its ShapeCollection, checks each shape's Width and Height against configurable pixel limits, logs any oversized shapes, optionally reduces them to the maximum allowed size, and saves the result to a new file.
// Keywords: Aspose.Cells shape validation | C# resize Excel shapes | shape width limit Aspose | shape height limit Aspose | Excel shape size enforcement | bulk shape resizing .NET | programmatic shape dimension check
// Common Searches: how to limit shape size in Excel with Aspose.Cells | C# script to find and resize oversized shapes in a workbook | Aspose.Cells iterate shapes and enforce width height constraints | validate Excel shape dimensions programmatically
// Developer Intent: Identify shapes that exceed defined width or height thresholds in an Excel file and automatically adjust them to comply with the limits.
// Use Cases: Audit and correct shape sizes before distributing a report to maintain layout consistency. | Automatically shrink images, charts, or SmartArt that violate corporate formatting standards. | Process large batches of workbooks to enforce uniform shape dimensions across all files.
// AI Prompts: Create a reusable C# method that receives a workbook path, max width, and max height, then returns a list of shapes that exceed those dimensions. | Show how to preserve aspect ratio while scaling down oversized shapes using Aspose.Cells. | Write unit tests for shape size validation that mock a workbook containing shapes of various widths and heights.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeSizeValidator
{
    // Loads an Excel workbook, walks through every worksheet and its ShapeCollection, checks each shape's Width and Height against configurable pixel limits, logs any oversized shapes, optionally reduces them to the maximum allowed size, and saves the result to a new file.
    class Program
    {
        // Define maximum allowed dimensions (in pixels)
        const int MaxWidth = 300;   // example: 300 pixels
        const int MaxHeight = 200;  // example: 200 pixels

        static void Main(string[] args)
        {
            // Input workbook path (change as needed or pass via command line)
            string inputPath = "input.xlsx";
            // Output workbook path (optional, if you want to save adjustments)
            string outputPath = "output_validated.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Examine each shape
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Check width and height against thresholds
                    bool widthExceeds = shape.Width > MaxWidth;
                    bool heightExceeds = shape.Height > MaxHeight;

                    if (widthExceeds || heightExceeds)
                    {
                        Console.WriteLine($"Worksheet '{sheet.Name}', Shape #{i} (Type={shape.Type}, Name={shape.Name}) exceeds limits.");

                        // Report current dimensions
                        Console.WriteLine($"  Current Size -> Width: {shape.Width}px, Height: {shape.Height}px");

                        // Optionally adjust the shape to fit within limits
                        if (widthExceeds)
                        {
                            shape.Width = MaxWidth;
                            Console.WriteLine($"  Width adjusted to {MaxWidth}px");
                        }

                        if (heightExceeds)
                        {
                            shape.Height = MaxHeight;
                            Console.WriteLine($"  Height adjusted to {MaxHeight}px");
                        }
                    }
                }
            }

            // Save the workbook (if any adjustments were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Validation complete. Workbook saved as '{outputPath}'.");
        }
    }
}
