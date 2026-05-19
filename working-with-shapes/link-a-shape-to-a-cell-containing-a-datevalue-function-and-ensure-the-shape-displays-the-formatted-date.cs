using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell that will hold the DATEVALUE formula
            // For example, use cell B2
            Cell targetCell = sheet.Cells["B2"];
            // Set the formula to DATEVALUE with a sample date string
            targetCell.Formula = "=DATEVALUE(\"2023-08-01\")";
            // Apply a date format to the cell so the value is displayed as a date
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format (e.g., mm/dd/yyyy)
            targetCell.SetStyle(dateStyle);

            // Add a rectangle shape that will be linked to the cell
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 120, 30);
            // Link the shape to the cell containing the DATEVALUE formula
            shape.LinkedCell = "$B$2";

            // Ensure the shape reflects the current value of the linked cell
            shape.UpdateSelectedValue();

            // Optionally, set the shape's text to display the linked value
            // The shape automatically shows the linked value; this line is just for clarity
            shape.Text = ""; // Clear any existing text; the linked value will appear

            // Save the workbook to a file
            workbook.Save("ShapeLinkedToDateValue.xlsx");
        }
    }
}