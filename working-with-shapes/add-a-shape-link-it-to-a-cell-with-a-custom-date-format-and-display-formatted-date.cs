using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedDateExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a date value in cell B2
        Cell dateCell = sheet.Cells["B2"];
        dateCell.PutValue(DateTime.Now);

        // Apply a custom date format to the cell (e.g., "dd-mmm-yyyy")
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "dd-mmm-yyyy";
        dateCell.SetStyle(dateStyle);

        // Add a label shape that will display the linked cell value
        // Parameters: upper left row, upper left column, top, left, height, width
        // Here we place the shape near cell B2
        Label label = (Label)sheet.Shapes.AddLabel(1, 1, 0, 0, 100, 30);
        label.Text = ""; // Text will be taken from the linked cell

        // Link the label shape to cell B2
        // false = not R1C1 style, true = use local (A1) notation
        label.SetLinkedCell("$B$2", false, true);

        // Ensure the shape updates its displayed value
        label.UpdateSelectedValue();

        // Save the workbook
        workbook.Save("ShapeLinkedDate.xlsx");
    }
}