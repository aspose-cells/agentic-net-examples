using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroAssignment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable macros in the workbook settings
            workbook.Settings.EnableMacros = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Form control button to the worksheet
            // Parameters: upper left row, upper left column, upper left pixel row, upper left pixel column, width, height
            Button button = sheet.Shapes.AddButton(2, 2, 0, 0, 100, 30);

            // Set the macro name that will be executed when the button is clicked
            button.MacroName = "MyMacro()";

            // Optionally set a caption for the button
            button.Text = "Run Macro";

            // Save the workbook as a macro‑enabled file
            workbook.Save("FormControlWithMacro.xlsm", SaveFormat.Xlsm);
        }
    }
}