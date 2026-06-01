using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate cells used in the nested formula
            cells["A1"].PutValue(5);   // First operand
            cells["A2"].PutValue(3);   // Second operand inside SUM
            cells["B1"].PutValue(2);   // Multiplied inside SUM

            // Set a nested formula in C1: =A1 + SUM(A2, B1*2)
            // Use the Formula property (compatible with all Aspose.Cells versions)
            cells["C1"].Formula = "=A1 + SUM(A2, B1*2)";

            // Calculate all formulas in the workbook
            wb.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, height, width, lowerRightRow, lowerRightColumn
            Shape rect = ws.Shapes.AddRectangle(1, 1, 100, 50, 2, 2);

            // Link the shape to the cell containing the nested formula (C1)
            rect.SetLinkedCell("$C$1", false, false);

            // Verify the linked cell address
            Console.WriteLine("Shape linked to cell: " + rect.LinkedCell);

            // Update the shape's selected value based on the linked cell (if applicable)
            ws.Shapes.UpdateSelectedValue();

            // Retrieve and display the calculated result from the linked cell
            Console.WriteLine("Calculated value in C1: " + cells["C1"].Value);
            // Expected calculation: 5 + SUM(3, 2*2) = 5 + (3 + 4) = 12
            Console.WriteLine("Expected value: 12");

            // Save the workbook (optional verification)
            string outputPath = "ShapeLinkedCellDemo.xlsx";
            try
            {
                wb.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}