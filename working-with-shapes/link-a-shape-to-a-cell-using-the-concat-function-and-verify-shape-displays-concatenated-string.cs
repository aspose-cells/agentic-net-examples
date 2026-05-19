using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkedCellDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells that will be concatenated
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set a formula in C1 that concatenates A1 and B1
            sheet.Cells["C1"].Formula = "=CONCAT(A1,B1)";

            // Add a rectangle shape to the worksheet
            Shape shape = sheet.Shapes.AddRectangle(2, 2, 200, 50, 0, 0);

            // Link the shape to the cell containing the concatenated result
            shape.LinkedCell = "C1";

            // Ensure the shape displays the latest value from the linked cell
            shape.UpdateSelectedValue();

            // Verify: output the linked cell address and its evaluated value
            string linkedAddress = shape.GetLinkedCell(false, false); // absolute A1 style
            string linkedValue = sheet.Cells[linkedAddress].StringValue;

            Console.WriteLine($"Shape is linked to cell: {linkedAddress}");
            Console.WriteLine($"Evaluated value in linked cell: {linkedValue}");

            // Save the workbook (optional, demonstrates lifecycle rule)
            workbook.Save("ShapeLinkedCellDemo.xlsx");
        }
    }
}