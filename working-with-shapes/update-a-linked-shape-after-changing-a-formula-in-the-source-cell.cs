using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkedShapeUpdate
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source cells with sample data
            // A1 will hold the value that the shape is linked to
            sheet.Cells["A1"].Value = 10;

            // Add a ListBox shape (dropdown) to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            Shape listBoxShape = sheet.Shapes.AddListBox(2, 0, 2, 0, 130, 130);

            // Set the input range for the ListBox (optional, just for demonstration)
            listBoxShape.SetInputRange("$A$1:$A$5", false, false);

            // Link the ListBox's selected value to cell A1
            listBoxShape.SetLinkedCell("$A$1", false, true);

            // Initial update – the shape reads the current value from A1 (10)
            listBoxShape.UpdateSelectedValue();

            // Verify the initial selection (should correspond to value 10)
            Console.WriteLine("Initial linked cell value: " + sheet.Cells["A1"].Value);

            // Change the formula/value in the source cell (A1)
            // For example, set a formula that calculates a new value
            sheet.Cells["A1"].Formula = "=SUM(5, 7)"; // Result will be 12

            // Recalculate the workbook so the formula result is materialized
            workbook.CalculateFormula();

            // Update the shape to reflect the new linked cell value
            listBoxShape.UpdateSelectedValue();

            // Output the updated value to verify the shape has been refreshed
            Console.WriteLine("Updated linked cell value after formula change: " + sheet.Cells["A1"].Value);

            // Save the workbook (lifecycle rule)
            workbook.Save("LinkedShapeUpdateDemo.xlsx");
        }
    }
}