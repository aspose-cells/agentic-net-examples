using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Input and output file paths
            string inputFile = "input.xlsx";
            string outputFile = "validated_output.xlsx";

            // Define maximum allowed dimensions (in pixels)
            const int MaxWidth = 500;
            const int MaxHeight = 300;

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputFile);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Examine each shape
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Check width and height against thresholds
                    bool exceedsWidth = shape.Width > MaxWidth;
                    bool exceedsHeight = shape.Height > MaxHeight;

                    if (exceedsWidth || exceedsHeight)
                    {
                        Console.WriteLine($"Worksheet '{sheet.Name}', Shape #{i} (Type={shape.Type}) exceeds limits:");
                        Console.WriteLine($"   Width: {shape.Width} (max {MaxWidth})");
                        Console.WriteLine($"   Height: {shape.Height} (max {MaxHeight})");

                        // Optionally resize the shape to fit within limits
                        if (exceedsWidth) shape.Width = MaxWidth;
                        if (exceedsHeight) shape.Height = MaxHeight;
                    }
                }
            }

            // Save the workbook (uses the provided save rule)
            workbook.Save(outputFile, SaveFormat.Xlsx);
        }
    }
}