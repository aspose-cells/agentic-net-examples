using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroButtonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable macros in the workbook (required for .xlsm files)
            workbook.Settings.EnableMacros = true;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a button to the worksheet (row 2, column 2, size 100x30 pixels)
            // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
            Button button = sheet.Shapes.AddButton(1, 0, 1, 0, 30, 100);

            // Set button display text
            button.Text = "Calculate Totals";

            // Assign the macro name to the button
            // The macro must exist in the VBA project; here we just set the reference.
            button.MacroName = "CalculateTotals";

            // Save the workbook as a macro‑enabled file
            workbook.Save("ButtonWithMacro.xlsm", SaveFormat.Xlsm);
        }
    }
}