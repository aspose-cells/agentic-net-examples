using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (row, column, upperLeftPixel, upperTopPixel, width, height)
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 100, 50, 200, 100);

            // Assign a macro name to the shape.
            // The macro "CalculateSummaryStats" should exist in the workbook's VBA project.
            shape.MacroName = "CalculateSummaryStats()";

            // Optional: give the shape a visible name
            shape.Name = "StatsButton";

            // Save the workbook (the macro will be linked to the shape)
            workbook.Save("WorkbookWithMacroShape.xlsx", SaveFormat.Xlsx);
        }
    }
}