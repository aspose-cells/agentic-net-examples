using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroAssignment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape that will act as a form control (e.g., a button)
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset,
            // lower right row, lower right column
            Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 4, 4);

            // Assign a macro name to the shape.
            // The macro must exist in the VBA project; here we just set the reference.
            shape.MacroName = "MyMacro()";

            // Optionally give the shape a visible caption
            shape.Text = "Run Macro";

            // Save the workbook as a macro‑enabled file (XLSM)
            workbook.Save("MacroAssignedShape.xlsm", SaveFormat.Xlsm);

            Console.WriteLine("Workbook saved with shape macro assignment.");
        }
    }
}