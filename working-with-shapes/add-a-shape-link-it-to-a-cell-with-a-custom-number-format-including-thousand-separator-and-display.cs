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
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value into cell B2
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue(1234567);

            // Create a custom style with thousand separator format "#,##0"
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "#,##0";

            // Apply the custom style to the target cell
            targetCell.SetStyle(customStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, upper left x offset, upper left y offset, width, height
            Shape rectShape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 4, 0, 0, 0, 150, 50);

            // Link the shape's value to cell B2 (absolute reference, A1 style, locale-aware)
            rectShape.SetLinkedCell("$B$2", false, true);

            // Optionally set some text for the shape (will display the linked cell's formatted value)
            rectShape.Text = "Linked Value";

            // Save the workbook
            workbook.Save("ShapeLinkedWithCustomNumberFormat.xlsx");
        }
    }
}