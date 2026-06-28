using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (row, column, upperLeftPixel, upperLeftPixel, width, height)
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 100, 100, 200, 100);

            // Assign a macro that calculates summary statistics to the shape
            // The macro must exist in the workbook's VBA project (e.g., in a standard module)
            shape.MacroName = "CalculateSummaryStats()";

            // Optional: give the shape a visible text label
            shape.Text = "Run Summary Stats";

            // Save the workbook (the macro will be linked to the shape)
            workbook.Save("MacroShapeExample.xlsm", SaveFormat.Xlsm);
        }
    }
}