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

            // Set a numeric value in cell A1
            Cell targetCell = worksheet.Cells["A1"];
            targetCell.PutValue(1234.56);

            // Apply a custom number format for currency (e.g., $1,234.56)
            Style style = targetCell.GetStyle();
            style.Custom = "$#,##0.00";
            targetCell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top, left, height, width (in points)
            RectangleShape rectShape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 50);

            // Optionally set some text for the shape (not required for linking)
            // rectShape.Text = "Currency Value";

            // Link the shape to the cell A1 using the SetLinkedCell method
            // Parameters: formula, isR1C1, isLocal
            rectShape.SetLinkedCell("$A$1", false, true);

            // Verify the link by reading the LinkedCell property
            Console.WriteLine("Shape is linked to cell: " + rectShape.LinkedCell);

            // Verify the displayed value of the linked cell (formatted as currency)
            Console.WriteLine("Linked cell formatted value: " + targetCell.StringValue);

            // Save the workbook to a file
            workbook.Save("ShapeLinkedCellCurrency.xlsx");
        }
    }
}