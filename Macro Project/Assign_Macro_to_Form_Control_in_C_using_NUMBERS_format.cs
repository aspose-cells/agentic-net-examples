using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsMacroDemo
{
    public class AssignMacroToFormControl
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a Forms button control to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, width, height
            Shape buttonShape = sheet.Shapes.AddButton(2, 2, 0, 0, 100, 30);

            // Cast the shape to Button (inherits from Shape)
            Button button = (Button)buttonShape;

            // Assign a macro to the button (macro name must include parentheses)
            button.MacroName = "MyMacro()";

            // Link the button to cell A1 (isR1C1 = false, isAbsolute = false)
            button.SetLinkedCell("A1", false, false);

            // Set a numeric value in the linked cell
            Cell linkedCell = sheet.Cells["A1"];
            linkedCell.PutValue(1234.5678);

            // Apply a number format (2 = "0.00")
            Style style = linkedCell.GetStyle();
            style.Number = 2;
            linkedCell.SetStyle(style);

            // Save the workbook
            workbook.Save("MacroAssignedButton.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            AssignMacroToFormControl.Run();
        }
    }
}