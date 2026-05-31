using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape (row, column, upper‑left pixel row, upper‑left pixel column, height, width)
                Shape shape = worksheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 100);

                // Move the shape to the back (lowest Z‑order) safely.
                // The shape is newly added, its ZOrderPosition is 0, so no movement is needed.
                // If the shape had a higher Z‑order, we would move it back by the required steps.
                int stepsToBack = -shape.ZOrderPosition;
                if (stepsToBack != 0)
                {
                    shape.ToFrontOrBack(stepsToBack);
                }

                // Output the current Z‑order position
                Console.WriteLine("Shape ZOrderPosition: " + shape.ZOrderPosition);

                // Save the workbook
                string outputPath = "ShapeZOrder.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}