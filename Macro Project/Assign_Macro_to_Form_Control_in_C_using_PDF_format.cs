using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroToFormControlPdf
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable macros in the workbook (required to assign a macro to a control)
        workbook.Settings.EnableMacros = true;

        // Add a Forms button to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Button button = worksheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

        // Set the macro name that will be invoked when the button is clicked
        button.MacroName = "MyMacro";

        // Optional: set the button caption
        button.Text = "Run Macro";

        // Save the workbook as a PDF file
        workbook.Save("ButtonWithMacro.pdf", SaveFormat.Pdf);
    }
}