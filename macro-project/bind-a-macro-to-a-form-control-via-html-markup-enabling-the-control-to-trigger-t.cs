using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroBindingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Forms button shape to the worksheet (row, column, upper-left offset, lower-right offset, width, height)
            // Here we place it at cell B2 with a size of 100x30 pixels
            Button button = (Button)sheet.Shapes.AddButton(1, 1, 0, 0, 100, 30);

            // Set the display text of the button
            button.Text = "Run Macro";

            // Bind a VBA macro to the button using the MacroName property
            // The macro name should match a macro defined in the workbook (e.g., Sub DoWork())
            button.MacroName = "DoWork()";

            // Optionally, embed HTML markup that also references the macro.
            // This demonstrates how the control can be triggered via HTML if the workbook is opened in a web context.
            button.HtmlText = "<button onclick=\"DoWork()\">Run Macro</button>";

            // Save the workbook (the macro itself must be added separately via VBA or other means)
            workbook.Save("MacroBoundButton.xlsx");
        }
    }
}