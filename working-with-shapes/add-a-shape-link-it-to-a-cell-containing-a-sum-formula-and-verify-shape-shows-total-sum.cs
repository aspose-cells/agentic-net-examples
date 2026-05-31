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

            // Populate data in column A (A1:A4)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);

            // Set a SUM formula in cell B5 that sums A1:A4
            sheet.Cells["B5"].Formula = "=SUM(A1:A4)";

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offset (pixels),
            // lower right row, lower right column, lower right offset (pixels)
            Shape rect = sheet.Shapes.AddRectangle(5, 1, 0, 5, 3, 0);

            // Link the shape to the cell containing the SUM formula (B5)
            rect.LinkedCell = "B5";

            // Update the shape's displayed value based on the linked cell
            rect.UpdateSelectedValue();

            // Retrieve the value from the linked cell
            double sumValue = sheet.Cells[rect.LinkedCell].DoubleValue;

            // Output the verification result
            Console.WriteLine($"Linked cell: {rect.LinkedCell}");
            Console.WriteLine($"Sum calculated in linked cell: {sumValue}");

            // Save the workbook
            string outputPath = "ShapeLinkedCellOutput.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}