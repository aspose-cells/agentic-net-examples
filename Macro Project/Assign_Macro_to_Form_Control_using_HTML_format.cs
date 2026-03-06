using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class AssignMacroToFormControl
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a Forms button (Button shape) to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        Button button = (Button)worksheet.Shapes.AddButton(2, 2, 0, 0, 100, 30);

        // Set the visible caption of the button
        button.Text = "Run Macro";

        // Assign the macro that will be executed when the button is clicked
        button.MacroName = "MyMacro()";

        // (Optional) Set HTML text for the button – this can be used for richer formatting
        button.HtmlText = "<a href=\"macro://MyMacro()\">Run Macro</a>";

        // Save the workbook to a file
        workbook.Save("MacroButton.xlsx");
    }
}