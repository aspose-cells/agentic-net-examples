using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class UngroupShapesDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputFile = "GroupedShapes.xlsx";
            const string outputFile = "UngroupedShapesResult.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Load the workbook containing grouped shapes
                Workbook workbook = new Workbook(inputFile);
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of shapes on the worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Iterate in reverse to avoid index issues when ungrouping
                for (int i = shapes.Count - 1; i >= 0; i--)
                {
                    Shape shape = shapes[i];

                    // Process only group shapes
                    if (shape.IsGroup)
                    {
                        GroupShape groupShape = (GroupShape)shape;
                        groupShape.Ungroup(); // Ungroup while preserving formatting
                    }
                }

                // Save the workbook with ungrouped shapes
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Processing failed: {ex.Message}");
            }
        }
    }
}