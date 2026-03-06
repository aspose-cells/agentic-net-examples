using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook (XLSX format)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape that will act as a form control (e.g., a button)
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        Shape shape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 30);

        // Assign a macro name to the shape. This macro must exist in the workbook's VBA project.
        shape.MacroName = "MyMacro()";

        // Optionally set a visible caption for the shape
        shape.Text = "Run Macro";

        // Save the workbook as an XLSX file (macros are not stored in XLSX, but the macro name is retained)
        workbook.Save("FormControlWithMacro.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Workbook saved with form control linked to macro.");
    }
}