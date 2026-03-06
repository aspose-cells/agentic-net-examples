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
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Forms button control to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            Button btn = (Button)sheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

            // Assign a macro to the button using the MacroName property
            // The macro name should match a VBA macro defined in the workbook (e.g., "MyMacro")
            btn.MacroName = "MyMacro";

            // Optionally set the button caption (display text)
            btn.Text = "Run Macro";

            // Save the workbook (as a macro-enabled file if you intend to add VBA later)
            workbook.Save("ButtonWithMacro.xlsx");
        }
    }
}