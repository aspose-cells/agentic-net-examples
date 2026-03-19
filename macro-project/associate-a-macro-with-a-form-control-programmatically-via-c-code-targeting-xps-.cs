using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will act as a Forms button
        // Parameters: upper left row, upper left column, width, height, offsetX, offsetY
        Shape buttonShape = worksheet.Shapes.AddRectangle(2, 2, 120, 30, 0, 0);

        // Assign a macro name to the shape (the macro must exist in the workbook's VBA project)
        buttonShape.MacroName = "MyMacro()";

        // Optionally set a caption to make it look like a button
        buttonShape.Text = "Run Macro";

        // Save the workbook as XPS
        workbook.Save("FormControlWithMacro.xps", SaveFormat.Xps);
    }
}