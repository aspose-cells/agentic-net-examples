using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled format will be used on save)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape that will act as a form control (button)
            // Parameters: upper left row, upper left column, width, height, top offset, left offset
            Shape shape = sheet.Shapes.AddRectangle(1, 1, 100, 30, 0, 0);

            // Assign a macro to the shape. The macro must exist in the VBA project of the workbook.
            shape.MacroName = "MyMacro()";

            // Optionally set a visible caption for the shape
            shape.Text = "Run Macro";

            // Save the workbook in macro‑enabled format (XLSM)
            workbook.Save("MacroButton.xlsm", SaveFormat.Xlsm);
        }
    }
}