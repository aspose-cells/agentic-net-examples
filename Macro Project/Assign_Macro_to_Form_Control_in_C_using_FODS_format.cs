using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (in memory)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a Forms button control to the worksheet.
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Button button = worksheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

        // Assign the macro name that should be executed when the button is clicked.
        // The macro must exist in the VBA project; here we just set the reference.
        button.MacroName = "MyMacro";

        // Save the workbook in FODS (OpenDocument Spreadsheet) format.
        // FODS supports macros when the workbook contains a VBA project.
        workbook.Save("ButtonWithMacro.fods", SaveFormat.Fods);
    }
}