using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a cell with text format and a sample value
            Cell linkedCell = sheet.Cells["B2"];
            linkedCell.PutValue("Formatted Text");

            // Apply text format (NumberFormat = "@")
            Style textStyle = workbook.CreateStyle();
            textStyle.Number = 49; // Text format
            linkedCell.SetStyle(textStyle);

            // Add a label shape to the worksheet (row, column, top, left, height, width)
            Label label = (Label)sheet.Shapes.AddLabel(2, 2, 0, 0, 100, 30);

            // Link the shape to the cell (use cell name without $ signs)
            label.LinkedCell = "B2";

            // Update the shape so it reflects the linked cell's value
            label.UpdateSelectedValue();

            // Verify: read back the linked cell address and its value
            Console.WriteLine("Shape's LinkedCell: " + label.LinkedCell);
            Console.WriteLine("Linked Cell Value: " + sheet.Cells[label.LinkedCell].StringValue);

            // Save the workbook (optional, just to persist the result)
            string outputPath = "ShapeLinkedCellDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}