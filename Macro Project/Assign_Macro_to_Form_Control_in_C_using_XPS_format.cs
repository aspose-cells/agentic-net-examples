using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsMacroToFormControl
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Form control button to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            Button button = sheet.Shapes.AddButton(1, 1, 0, 0, 120, 30);

            // Assign a macro name to the button (the macro must exist in the workbook's VBA project)
            button.MacroName = "MyMacro()";

            // Optional: set additional visual properties
            button.Name = "MyButton";
            button.Text = "Run Macro";

            // Prepare XPS save options (lifecycle: create)
            XpsSaveOptions saveOptions = new XpsSaveOptions
            {
                OnePagePerSheet = true,
                DefaultFont = "Arial"
            };

            // Save the workbook as XPS (lifecycle: save)
            workbook.Save("ButtonWithMacro.xps", saveOptions);

            Console.WriteLine("Workbook saved as XPS with a button linked to macro.");
        }
    }
}