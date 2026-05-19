using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapePositionAnalysis
{
    public class ShapePositionRetriever
    {
        /// <summary>
        /// Loads a workbook, iterates through all worksheets and their shapes,
        /// and writes the absolute position (X, Y) and size (Width, Height) of each shape to the console.
        /// </summary>
        /// <param name="workbookPath">Full path to the Excel file to be analyzed.</param>
        public static void DisplayShapePositions(string workbookPath)
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(workbookPath);

            // Iterate over each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = worksheet.Shapes;

                // Loop through all shapes in the collection
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Output shape details: worksheet name, shape index, name, type, and absolute position/size
                    Console.WriteLine(
                        $"Worksheet: {worksheet.Name}, " +
                        $"Shape Index: {i}, " +
                        $"Name: {shape.Name}, " +
                        $"Type: {shape.Type}, " +
                        $"X: {shape.X}, Y: {shape.Y}, " +
                        $"Width: {shape.Width}, Height: {shape.Height}, " +
                        $"ZOrder: {shape.ZOrderPosition}");
                }
            }
        }

        // Example usage
        public static void Main()
        {
            // Path to the workbook to analyze
            string path = "InputWorkbook.xlsx";

            // Display positions of all shapes
            DisplayShapePositions(path);
        }
    }
}