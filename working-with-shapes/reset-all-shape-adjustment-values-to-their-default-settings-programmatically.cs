using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ResetShapeAdjustments
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through every worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Iterate through every shape on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Access the geometry of the shape
                Geometry geometry = shape.Geometry;

                // If the shape has adjustment guides, reset each one
                if (geometry != null && geometry.ShapeAdjustValues.Count > 0)
                {
                    foreach (ShapeGuide guide in geometry.ShapeAdjustValues)
                    {
                        // Reset the guide value to its default (commonly 0)
                        guide.Value = 0;
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}