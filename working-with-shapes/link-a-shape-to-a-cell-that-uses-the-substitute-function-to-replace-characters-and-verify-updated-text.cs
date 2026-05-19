using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLinkShapeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a TextBox shape (acts as a form control)
            // Parameters: upper left row, upper left column, lower right row, lower right column, upper left offset, lower right offset
            Shape textBox = sheet.Shapes.AddTextBox(2, 1, 4, 3, 0, 0);

            // Link the shape to cell A1 (row 0, column 0)
            // The two boolean parameters indicate whether to use the cell's value as a formula and whether to use the cell's style.
            textBox.SetLinkedCell("A1", false, false);

            // Set a SUBSTITUTE formula in the linked cell.
            // This will replace "World" with "Aspose" in the string "Hello World".
            sheet.Cells["A1"].Formula = "SUBSTITUTE(\"Hello World\",\"World\",\"Aspose\")";

            // Refresh the shape so it reflects the current value of the linked cell
            textBox.UpdateSelectedValue();

            // Verify the updated text displayed by the shape
            Console.WriteLine("Shape text after linking and formula evaluation: " + textBox.Text);

            // Save the workbook to a file
            workbook.Save("LinkedShapeWithSubstitute.xlsx");
        }
    }
}