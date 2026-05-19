using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a raw numeric value in cell A1
            worksheet.Cells["A1"].PutValue(1234.567);

            // Use the TEXT function in cell B1 to format the number (e.g., two decimal places with a dollar sign)
            worksheet.Cells["B1"].Formula = @"=TEXT(A1,""$#,##0.00"")";

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels), upper left offset (pixels),
            // height (pixels), width (pixels)
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);

            // Link the shape to the formatted cell (B1). The shape will display the value of B1.
            shape.LinkedCell = "$B$1";

            // Optionally, update the shape's displayed value immediately
            shape.UpdateSelectedValue();

            // Save the workbook to a file
            workbook.Save("ShapeLinkedToTextFunction.xlsx");
        }
    }
}