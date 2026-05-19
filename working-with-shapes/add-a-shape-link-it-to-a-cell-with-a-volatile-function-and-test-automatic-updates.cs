using System;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkedCellDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet (row, column, width, height, top, left)
                Shape rectangle = worksheet.Shapes.AddRectangle(2, 2, 100, 50, 0, 0);

                // Link the shape to cell A1
                rectangle.SetLinkedCell("$A$1", false, true);

                // Set a volatile function (NOW) in the linked cell
                worksheet.Cells["A1"].Formula = "=NOW()";

                // Initial calculation to evaluate the volatile function
                workbook.CalculateFormula();

                // Output the initial value of the linked cell
                Console.WriteLine("Initial linked cell value (A1): " + worksheet.Cells["A1"].Value);

                // Simulate a delay to allow the volatile function to change
                Thread.Sleep(3000);

                // Recalculate the workbook to refresh the volatile function
                workbook.CalculateFormula();

                // Output the updated value of the linked cell
                Console.WriteLine("Updated linked cell value (A1): " + worksheet.Cells["A1"].Value);

                // Verify that the shape still points to the same cell
                Console.WriteLine("Shape's LinkedCell reference: " + rectangle.LinkedCell);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}